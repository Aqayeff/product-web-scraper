import csv
import re
import time
from pathlib import Path
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment


# ==============================
# SETTINGS
# ==============================

BASE_DIR = Path(__file__).resolve().parent

OUTPUT_DIR = BASE_DIR / "output"
CSV_FILE = OUTPUT_DIR / "books.csv"
EXCEL_FILE = OUTPUT_DIR / "books.xlsx"

BASE_URL = "https://books.toscrape.com/"

HEADERS = {
    "User-Agent": "Mozilla/5.0"
}

REQUEST_TIMEOUT = 10
DELAY_BETWEEN_PAGES = 1


# ==============================
# REQUEST
# ==============================

def get_page(url):
    response = requests.get(
        url,
        headers=HEADERS,
        timeout=REQUEST_TIMEOUT
    )

    response.raise_for_status()

    return response.text


# ==============================
# CLEAN PRICE
# ==============================

def clean_price(price_text):
    if not price_text:
        return None

    # Ищем число внутри текста
    match = re.search(r"\d+(?:\.\d+)?", price_text)

    if not match:
        return None

    try:
        return float(match.group())
    except ValueError:
        return None


# ==============================
# CLEAN RATING
# ==============================

def clean_rating(rating_text):
    ratings = {
        "One": 1,
        "Two": 2,
        "Three": 3,
        "Four": 4,
        "Five": 5
    }

    return ratings.get(rating_text, None)


# ==============================
# PARSE BOOKS
# ==============================

def parse_books(html):
    soup = BeautifulSoup(html, "html.parser")

    books = []

    book_cards = soup.select("article.product_pod")

    for book in book_cards:

        # --------------------------
        # TITLE
        # --------------------------

        title_element = book.select_one("h3 a")

        title = ""

        if title_element:
            title = title_element.get("title", "").strip()


        # --------------------------
        # PRICE
        # --------------------------

        price_element = book.select_one(".price_color")

        price = None

        if price_element:
            price_text = price_element.get_text(strip=True)
            price = clean_price(price_text)


        # --------------------------
        # RATING
        # --------------------------

        rating_element = book.select_one("p.star-rating")

        rating = None

        if rating_element:

            rating_classes = rating_element.get("class", [])

            if len(rating_classes) > 1:
                rating = clean_rating(rating_classes[1])


        # --------------------------
        # AVAILABILITY
        # --------------------------

        availability_element = book.select_one(".availability")

        availability = ""

        if availability_element:
            availability = availability_element.get_text(
                " ",
                strip=True
            )


        # --------------------------
        # URL
        # --------------------------

        product_url = ""

        if title_element:

            relative_url = title_element.get("href", "")

            if relative_url:
                product_url = urljoin(
                    BASE_URL,
                    relative_url
                )


        # --------------------------
        # SAVE BOOK
        # --------------------------

        books.append({
            "title": title,
            "price": price,
            "rating": rating,
            "availability": availability,
            "url": product_url
        })


    return books


# ==============================
# NEXT PAGE
# ==============================

def get_next_page(html, current_url):

    soup = BeautifulSoup(html, "html.parser")

    next_button = soup.select_one("li.next a")

    if not next_button:
        return None

    next_url = next_button.get("href", "")

    if not next_url:
        return None

    return urljoin(
        current_url,
        next_url
    )


# ==============================
# SCRAPE ALL PAGES
# ==============================

def scrape_all_pages():

    all_books = []

    current_url = BASE_URL

    page_number = 1


    while current_url:

        print(
            f"Scraping page {page_number}: {current_url}"
        )


        try:

            html = get_page(current_url)

            books = parse_books(html)

            print(
                f"  Found {len(books)} books"
            )

            all_books.extend(books)


            next_url = get_next_page(
                html,
                current_url
            )

            current_url = next_url

            page_number += 1


            if current_url:
                time.sleep(
                    DELAY_BETWEEN_PAGES
                )


        except requests.RequestException as error:

            print(
                f"  Request error: {error}"
            )

            break


    return all_books


# ==============================
# SAVE CSV
# ==============================

def save_csv(books, filename):

    fieldnames = [
        "title",
        "price",
        "rating",
        "availability",
        "url"
    ]


    with open(
        filename,
        "w",
        encoding="utf-8",
        newline=""
    ) as file:

        writer = csv.DictWriter(
            file,
            fieldnames=fieldnames
        )

        writer.writeheader()

        writer.writerows(books)


# ==============================
# SAVE EXCEL
# ==============================

def save_excel(books, filename):

    workbook = Workbook()

    sheet = workbook.active

    sheet.title = "Books"


    headers = [
        "Title",
        "Price",
        "Rating",
        "Availability",
        "URL"
    ]


    sheet.append(headers)


    for book in books:

        sheet.append([
            book["title"],
            book["price"],
            book["rating"],
            book["availability"],
            book["url"]
        ])


    # --------------------------
    # HEADER STYLE
    # --------------------------

    header_fill = PatternFill(
        "solid",
        fgColor="1F4E78"
    )

    header_font = Font(
        color="FFFFFF",
        bold=True
    )


    for cell in sheet[1]:

        cell.fill = header_fill

        cell.font = header_font

        cell.alignment = Alignment(
            horizontal="center"
        )


    # --------------------------
    # NUMBER FORMATS
    # --------------------------

    for cell in sheet["B"][1:]:

        cell.number_format = '£0.00'


    for cell in sheet["C"][1:]:

        cell.number_format = '0'


    # --------------------------
    # EXCEL SETTINGS
    # --------------------------

    sheet.freeze_panes = "A2"

    sheet.auto_filter.ref = sheet.dimensions


    sheet.column_dimensions["A"].width = 50
    sheet.column_dimensions["B"].width = 15
    sheet.column_dimensions["C"].width = 12
    sheet.column_dimensions["D"].width = 20
    sheet.column_dimensions["E"].width = 80


    # --------------------------
    # SAVE
    # --------------------------

    workbook.save(filename)


# ==============================
# MAIN
# ==============================

def main():

    OUTPUT_DIR.mkdir(
        exist_ok=True
    )


    print()
    print("==============================")
    print("PRODUCT WEB SCRAPER")
    print("==============================")
    print()


    books = scrape_all_pages()


    print()
    print(
        f"Total books collected: {len(books)}"
    )


    if not books:

        print(
            "No books were collected."
        )

        return


    print()
    print("Saving CSV...")

    save_csv(
        books,
        CSV_FILE
    )


    print("Saving Excel...")

    save_excel(
        books,
        EXCEL_FILE
    )


    print()
    print("==============================")
    print("SCRAPING COMPLETED")
    print("==============================")

    print(
        f"Books: {len(books)}"
    )

    print(
        f"CSV: {CSV_FILE}"
    )

    print(
        f"Excel: {EXCEL_FILE}"
    )

    print("==============================")


# ==============================
# START PROGRAM
# ==============================

if __name__ == "__main__":
    main()