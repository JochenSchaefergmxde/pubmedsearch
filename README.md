# PubMed Search and Visualization

This script is a command-line tool for searching the PubMed database, saving the results in various formats, and generating an interactive plot of the findings.

## Installation

1.  Clone the repository:
    ```bash
    git clone https://github.com/JochenSchaefergmxde/search_pubmed2.py.git
    cd search_pubmed2.py
    ```

2.  Install the required dependencies:
    ```bash
    pip install -r requirements.txt
    ```

## Usage

The script is run from the command line and accepts several arguments to customize the search and output.

```bash
python search_pubmed2.py -b <search_term> [-k <category>] [-m <max_count>] [--format <format>] [--plot]
```

### Arguments

*   `-b` (Required): The primary search term (e.g., "lyme", "borrelia").
*   `-k` (Optional): The category of items to search against. Defaults to `v` (Vitamins).
    *   `v`: Vitamins
    *   `c`: Cross-Reactivity
    *   `b`: Blutwerte (Blood Values)
    *   `p`: Pharma (Pharmaceuticals)
    *   `k`: Krankheiten (Diseases)
    *   `m`: Mineralien (Minerals)
    *   `a`: Aminosäuren (Amino Acids)
    *   `g`: Genetik (Genetics)
    *   `h`: Herb/Pflanzen (Herbs/Plants)
    *   `l`: Gute Bakterien (Good Bacteria)
    *   `e`: Enzyme
*   `-m` (Optional): The maximum number of publications to retrieve for each search query. Defaults to `20`.
*   `--format` (Optional): The output format for the results file. Defaults to `xls`.
    *   `xls`
    *   `csv`
    *   `json`
    *   `html`
*   `--plot` (Optional): If included, the script will generate an interactive HTML plot of the results, showing the number of studies found for each keyword.

### Example

To search for up to 10 publications about "lyme" in the "Vitamins" category, save the results as a CSV file, and generate a plot, you would run:

```bash
python search_pubmed2.py -b "lyme" -k "v" -m 10 --format csv --plot
```
This will create two files:
*   A CSV file with the search results (e.g., `20231027_123456_lyme_v.csv`).
*   An interactive HTML plot (e.g., `20231027_123456_lyme_v.html`).
