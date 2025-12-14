import argparse
import datetime
import inspect
import logging
import re
from urllib import parse

import pandas as pd
import xlwt
from Bio import Entrez, Medline

# Import search term lists from the configuration file
from config import (
    AMINOSAEUREN,
    BLUTWERTE,
    CODEWOERTER,
    ENZYME,
    GUTEBAKTERIEN,
    HAUPTGENETIK,
    KRANKHEITEN,
    MEDIKAMENTE,
    MINERALIEN,
    PFLANZEN,
    VITAMIN,
)

# Constants
LJUST = 30
MAX_LINE_LENGTH = 80

def setup_logging():
    """Sets up the logging configuration."""
    FORMAT = "[%(filename)s:%(lineno)s - %(funcName)20s() ] %(message)s"
    logging.basicConfig(filename='pubmed_bio.log', level=logging.DEBUG, format=FORMAT)
    logging.debug('Log file initialized.')

def autolog(message):
    """Automatically log the current function details."""
    message = sonderzeichen(message)
    func = inspect.currentframe().f_back.f_code
    logging.debug(f"{message}: {func.co_name} in {func.co_filename}:{func.co_firstlineno}")

def sonderzeichen(text):
    """Removes special characters and truncates the text."""
    if not isinstance(text, str):
        text = str(text)
    text = text.encode('ascii', 'ignore').decode('ascii')
    text = re.sub(r'["-]', ' ', text)
    text = re.sub(r'/', ' div ', text)
    return text[:200]


def fetch_pubmed_records(term, max_count):
    """Fetches PubMed records for a given search term."""
    print(f"Searching PubMed for: '{term}' (max: {max_count} records)...")
    Entrez.email = 'jochen_schaefer@gmx.de'
    try:
        search_handle = Entrez.esearch(db='pubmed', retmax=max_count, term=term, usehistory="y")
        search_results = Entrez.read(search_handle)
        search_handle.close()

        ids = search_results['IdList']
        print(f"Found {search_results['Count']} publications for '{term}'.")

        if not ids:
            return [], 0

        fetch_handle = Entrez.efetch(db='pubmed', id=ids, rettype='medline', retmode='text')
        records = Medline.parse(fetch_handle)
        return list(records), search_results['Count']

    except Exception as e:
        print(f"Error fetching from PubMed: {e}")
        autolog(f"PubMed fetch error: {e}")
        return [], 0

def process_record(record, search_term, search_category, total_count, max_count, item):
    """Processes a single PubMed record and returns a dictionary."""
    pmid = record.get("PMID", "?")
    title = sonderzeichen(record.get("TI", "?"))
    aid = record.get("AID", "?")

    click_url = f"http://www.ncbi.nlm.nih.gov/pubmed/?term={pmid}"
    encoded_term = parse.quote(search_term)
    search_url = f"http://www.ncbi.nlm.nih.gov/pubmed/?term={encoded_term}"

    scihub_url = f"http://sci-hub.ren/{aid[0]}" if isinstance(aid, list) and aid else ""

    codeword = ""
    keywords = {'deficien': 'deficiency', 'low': 'low', 'co-factor': 'Co-Factor', 'therapy': 'therapy'}
    for keyword, label in keywords.items():
        if keyword in title:
            codeword = label
            break

    return {
        "Anzahl": total_count,
        "Max Anzahl": max_count,
        "Bacterium": search_category,
        "Vitamin": item,
        "Suchstring": search_url,
        "PMIDLink US": click_url,
        "Titel": title,
        "Codewort": codeword,
        "PMID": pmid,
        "SCIHUB": scihub_url,
    }

def scann(bacterias, search_list, max_count, use_codeword, args):
    """Scans PubMed for publications and returns a list of results."""
    results = []
    for bacteria in bacterias:
        for item in search_list:
            term = f'({bacteria}[Title/Abstract] AND {item})'
            if args.k == "c":
                term = f'({bacteria}[Title/Abstract] AND {item} AND cross-reactivity[Title/Abstract])'
            elif use_codeword:
                term += ' AND (low OR deficiency OR therapy OR inhibitor)'

            records, total_count = fetch_pubmed_records(term, max_count)

            if not records:
                continue

            for record in records:
                processed_record = process_record(record, term, bacteria, total_count, max_count, item)
                results.append(processed_record)
    return results

def get_search_list(category_key):
    """Returns the appropriate search list based on the category key."""
    category_map = {
        "g": HAUPTGENETIK, "k": KRANKHEITEN, "c": KRANKHEITEN,
        "b": BLUTWERTE, "p": MEDIKAMENTE, "m": MINERALIEN,
        "a": AMINOSAEUREN, "h": PFLANZEN, "l": GUTEBAKTERIEN,
        "e": ENZYME, "v": VITAMIN
    }
    search_list = category_map.get(category_key, VITAMIN)
    use_codeword = (category_key == 'v')
    return search_list, use_codeword

def save_results(df, filename, file_format):
    """Saves the DataFrame to the specified file format."""
    print(f"Saving results to '{filename}' in {file_format} format...")
    try:
        if file_format == 'xls':
            df.to_excel(filename, index=False)
        elif file_format == 'csv':
            df.to_csv(filename, index=False)
        elif file_format == 'json':
            df.to_json(filename, orient='records', indent=4)
        elif file_format == 'html':
            df.to_html(filename, index=False)
        print("Successfully saved results.")
    except Exception as e:
        print(f"Error saving file: {e}")
        autolog(f"File save error: {e}")

def main():
    parser = argparse.ArgumentParser(description='Pubmed search and translation')
    parser.add_argument('-b', help="Krankheit nach der gesucht werden soll", required=True)
    parser.add_argument('-m', default=20, type=int, help="Maximale Anzahl der angezeigten Studien")
    parser.add_argument('-k', default="v",
                        help="suche gegen implementierte Datensätze \n c=Cross-Reactivity b=Blutwerte p=Pharma k=Krankheiten v=VITAMINE m=MINERALIEN a=AMINOSÄUREN g=Genetik h=Herb/Pflanzen l=Gute Bakterien e=Enzyme")
    parser.add_argument('--format', default='xls', choices=['xls', 'csv', 'json', 'html'],
                        help="Output format for the results.")
    args = parser.parse_args()

    setup_logging()

    search_list, use_codeword = get_search_list(args.k)

    results = scann([args.b], search_list, args.m, use_codeword, args)

    if not results:
        print("No results found.")
        return

    df = pd.DataFrame(results)

    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    sanitized_bacteria = sonderzeichen(args.b).replace(' ', '_')
    base_filename = f"{timestamp}_{sanitized_bacteria}_{args.k}"
    filename = f"{base_filename}.{args.format}"

    save_results(df, filename, args.format)

if __name__ == '__main__':
    main()