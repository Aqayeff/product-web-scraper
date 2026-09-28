# Product Web Scraper

A Python web scraper that collects product information from a website, processes the data, and exports the results to CSV and Excel files.

The project demonstrates practical web scraping, HTML parsing, pagination handling, data cleaning, error handling, and automated report generation.

## Features

* Scrapes product data from multiple pages
* Automatically follows pagination
* Collects:

  * Product title
  * Price
  * Rating
  * Availability
  * Product URL
* Cleans price data and converts it to numeric values
* Converts ratings from text to numbers
* Handles request errors
* Adds a delay between page requests
* Exports data to CSV
* Exports data to formatted Excel
* Freezes the Excel header row
* Adds Excel filters
* Applies column widths
* Formats prices as currency
* Uses reusable functions and configurable settings

## Technologies

* Python 3
* Requests
* BeautifulSoup
* OpenPyXL
* CSV
* pathlib
* urllib.parse

## Project Structure

```text
product-web-scraper/
│
├── data/
│
├── output/
│   ├── books.csv
│   └── books.xlsx
│
├── main.py
├── README.md
└── requirements.txt
```

## How It Works

The scraper follows this workflow:

```text
Website
   ↓
Send HTTP request
   ↓
Receive HTML
   ↓
Parse HTML with BeautifulSoup
   ↓
Extract product data
   ↓
Clean and normalize data
   ↓
Follow next page
   ↓
Repeat until the last page
   ↓
Save CSV
   ↓
Save Excel
```

## Data Collected

Each product contains:

| Field        | Description                              |
| ------------ | ---------------------------------------- |
| Title        | Product title                            |
| Price        | Product price as a numeric value         |
| Rating       | Rating converted to a number from 1 to 5 |
| Availability | Product availability status              |
| URL          | Product page URL                         |

## Pagination

The scraper automatically detects the **Next** page button and continues scraping until there are no more pages.

Relative page URLs are converted into complete URLs using `urljoin()`.

This allows the scraper to work across the entire catalog instead of collecting data from only one page.

## Data Cleaning

Price values are converted from text into numeric values.

For example:

```text
£51.77 → 51.77
```

Ratings are also converted from text into numbers:

```text
One   → 1
Two   → 2
Three → 3
Four  → 4
Five  → 5
```

This makes the exported data easier to analyze and process later.

## Output

The scraper creates two files inside the `output/` directory.

### CSV

```text
output/books.csv
```

The CSV file contains the raw structured product dataset.

### Excel

```text
output/books.xlsx
```

The Excel report includes:

* formatted headers
* filters
* frozen header row
* adjusted column widths
* numeric prices
* currency formatting
* numeric ratings

## Installation

Install the required packages:

```bash
pip install -r requirements.txt
```

Or install them manually:

```bash
pip install requests beautifulsoup4 openpyxl
```

## Running the Scraper

Run:

```bash
python main.py
```

The program will display the scraping progress in the terminal:

```text
PRODUCT WEB SCRAPER

Scraping page 1...
Scraping page 2...
Scraping page 3...
...
```

After completion, the CSV and Excel files will be available in the `output/` directory.

## Configuration

Several settings can be changed at the top of `main.py`:

```python
BASE_URL = "https://books.toscrape.com/"

REQUEST_TIMEOUT = 10

DELAY_BETWEEN_PAGES = 1
```

This makes the scraper easier to adapt to another website or project.

## Error Handling

The scraper uses request error handling to prevent the program from crashing when a page cannot be downloaded.

Example:

```python
try:
    html = get_page(current_url)
except requests.RequestException as error:
    print(f"Request error: {error}")
```

The scraper also checks whether products, prices, ratings, availability information, and URLs exist before extracting them.

## Example Use Cases

This type of scraper can be adapted for different data collection tasks, for example:

* Product catalogs
* E-commerce research
* Price monitoring
* Competitor research
* Real estate listings
* Job listings
* Public directories
* Market research datasets

The HTML selectors and output fields would need to be changed for each target website.

## Disclaimer

This project was created for educational and portfolio purposes using a public website specifically designed for web scraping practice.

When scraping real websites, always check the website's terms of service, robots.txt, rate limits, and applicable laws before collecting data.

## Author

Amil

Python / Data Processing / Web Scraping Portfolio
