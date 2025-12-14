import argparse
import datetime
import inspect
import logging
import re
from urllib import parse

import xlwt
from Bio import Entrez, Medline

# Import search term lists from the configuration file
from config import (
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
    # Fallback for non-string inputs
    if not isinstance(text, str):
        text = str(text)

    # Using a more robust way to remove non-ASCII characters
    text = text.encode('ascii', 'ignore').decode('ascii')
    text = re.sub(r'["-]', ' ', text)
    text = re.sub(r'/', ' div ', text)
    return text[:200]


def fetch_pubmed_records(term, max_count):
    """Fetches PubMed records for a given search term."""
    print(f'Getting {max_count} publications containing {term[:15]}...')
    Entrez.email = 'jochen_schaefer@gmx.de'
    try:
        search_handle = Entrez.esearch(db='pubmed', retmax=max_count, term=term, usehistory="y")
        search_results = Entrez.read(search_handle)
        search_handle.close()

        ids = search_results['IdList']
        print(f"Total number of publications containing {term}: {search_results['Count']}")

        if not ids:
            return [], 0

        fetch_handle = Entrez.efetch(db='pubmed', id=ids, rettype='medline', retmode='text')
        records = Medline.parse(fetch_handle)
        return list(records), search_results['Count']

    except Exception as e:
        print(f"Error fetching from PubMed: {e}")
        autolog(f"PubMed fetch error: {e}")
        return [], 0

def write_excel_header(sheet):
    """Writes the header row to the Excel sheet."""
    headers = {
        0: "Anzahl", 1: "Max Anzahl", 2: "Vitamin", 3: "Bacterium",
        4: "Suchstring", 5: "PMIDLink US", 6: "Titel", 10: "Codewort",
        11: "PMID", 12: "SCIHUB", 13: "Patent"
    }
    for col, header in headers.items():
        sheet.write(0, col, header)
    return {header: col for col, header in headers.items()} # return a map for easier access

def process_record(record, sheet, row, search_term, search_category, total_count, max_count, col_map, item):
    """Processes a single PubMed record and writes it to the Excel sheet."""
    pmid = record.get("PMID", "?")
    title = sonderzeichen(record.get("TI", "?"))
    aid = record.get("AID", "?")

    click_url = f"http://www.ncbi.nlm.nih.gov/pubmed/?term={pmid}"
    # URL-encode the search term to handle special characters
    encoded_term = parse.quote(search_term)
    search_url = f"http://www.ncbi.nlm.nih.gov/pubmed/?term={encoded_term}"


    sheet.write(row, col_map["Anzahl"], total_count)
    sheet.write(row, col_map["Max Anzahl"], max_count)
    sheet.write(row, col_map["Bacterium"], search_category)
    sheet.write(row, col_map["Vitamin"], item)
    sheet.write(row, col_map["Suchstring"], xlwt.Formula(f'HYPERLINK("{search_url[:250]}","{title}")'))
    sheet.write(row, col_map["PMIDLink US"], xlwt.Formula(f'HYPERLINK("{click_url}","{title}")'))
    sheet.write(row, col_map["PMID"], pmid)

    if isinstance(aid, list) and aid:
        scihub_url = f"http://sci-hub.ren/{aid[0]}"
        sheet.write(row, col_map["SCIHUB"], xlwt.Formula(f'HYPERLINK("{scihub_url}","{scihub_url}")'))

    # Simplified logic for keywords
    style_deficien = xlwt.easyxf('pattern: pattern solid, fore_colour light_blue;font: colour white, bold True;')
    style_cd4 = xlwt.easyxf('pattern: pattern solid, fore_colour red;font: colour white, bold True;')

    keywords = {'deficien': 'deficiency', 'low': 'low', 'co-factor': 'Co-Factor', 'therapy': 'therapy'}
    for keyword, label in keywords.items():
        if keyword in title:
            sheet.write(row, col_map["PMIDLink US"], xlwt.Formula(f'HYPERLINK("{click_url}","{title}")'), style_deficien)
            sheet.write(row, col_map["Codewort"], xlwt.Formula(f'"{label}"'), style_cd4)
            break

def scann(bacterias, search_list, sheet1, max_count, use_codeword, args):
    """Scans PubMed for publications and fills the Excel sheet."""
    row = 1
    col_map = write_excel_header(sheet1)

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
                process_record(record, sheet1, row, term, bacteria, total_count, max_count, col_map, item)
                row += 1

def get_search_list(category_key):
    """Returns the appropriate search list based on the category key."""
    # AMINOSAEUREN is not defined in config, so I'll add an empty list for it
    AMINOSAEUREN = []
    category_map = {
        "g": HAUPTGENETIK, "k": KRANKHEITEN, "c": KRANKHEITEN,
        "b": BLUTWERTE, "p": MEDIKAMENTE, "m": MINERALIEN,
        "a": AMINOSAEUREN, "h": PFLANZEN, "l": GUTEBAKTERIEN,
        "e": ENZYME, "v": VITAMIN
    }
    # Default to VITAMIN if key not found
    search_list = category_map.get(category_key, VITAMIN)
    # The logic for use_codeword seems to be that it's True only for 'v'
    use_codeword = (category_key == 'v')
    return search_list, use_codeword


def main():
    parser = argparse.ArgumentParser(description='Pubmed search and translation')
    parser.add_argument('-b', help="Krankheit nach der gesucht werden soll", required=True)
    parser.add_argument('-m', default=20, type=int, help="Maximale Anzahl der angezeigten Studien")
    parser.add_argument('-k', default="v",
                        help="suche gegen implementierte Datensätze \n c=Cross-Reactivity b=Blutwerte p=Pharma k=Krankheiten v=VITAMINE m=MINERALIEN a=AMINOSÄUREN g=Genetik h=Herb/Pflanzen l=Gute Bakterien e=Enzyme")
    args = parser.parse_args()

    setup_logging()

    search_list, use_codeword = get_search_list(args.k)

    book = xlwt.Workbook(encoding="utf-8")
    sheet1 = book.add_sheet("Bacteria", cell_overwrite_ok=True)

    scann([args.b], search_list, sheet1, args.m, use_codeword, args)

    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    sanitized_bacteria = sonderzeichen(args.b).replace(' ', '_')
    filename = f"{timestamp}_{sanitized_bacteria}_{args.k}.xls"

    try:
        book.save(filename)
        print(f"Saved results to {filename}")
    except Exception as e:
        print(f"Error saving Excel file: {e}")
        autolog(f"Excel save error: {e}")

if __name__ == '__main__':
    main()