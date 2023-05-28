import time
import requests
import openpyxl
from loguru import logger
from bs4 import BeautifulSoup, Tag

from helpers import (
    remove_brakets,
    create_excel_tamplate,
    get_company_data,
    HEADERS,
    get_company_urls,
    get_cik,
)


def get_rows(document: BeautifulSoup) -> list:
    table_1: Tag = document.select_one(
        'b:-soup-contains("Table I")'
    ).parent.parent.parent.parent
    table_1_data = table_1.select_one("tbody")
    if not table_1_data:
        return []
    table_1_data = table_1_data.select("tr")
    result = []
    for row in table_1_data:
        row = [column_raw.text for column_raw in row if column_raw.text not in ["\n"]]
        result.append(row)

    return result


def parse_document(url: str) -> list:
    res = requests.get(url, headers=HEADERS)
    if res.status_code != 200:
        logger.warning(f"Could not get document with url {url}")
        return
    document = BeautifulSoup(res.content, "html.parser")
    rows = get_rows(document)
    if not rows:
        logger.debug("No table I found")
        return
    result = []
    validation_errors_count = 0
    for row in rows:
        date = row[1].strip()
        transaction_code = remove_brakets(row[3].strip())
        shares_amount = remove_brakets(row[5].strip().replace(",", ""))
        if shares_amount:
            shares_amount = float(shares_amount)
        price = remove_brakets(row[7].strip())
        if price.startswith("$"):
            price = price[1:]
            price.replace(",", "")
            price = float(price)
        # if price == "(1)":
        #     continue
        total_amount: str = remove_brakets(row[8].strip())
        name = (
            document.select_one('span:-soup-contains("Reporting Person")')
            .find_next("table")
            .text
        )
        try:
            relationship = (
                document.select_one('span:-soup-contains("Relationship")')
                .find_next("table")
                .select("tr")[2]
                .select_one("span")
                .text
            )
        except:
            relationship = None
        ad = row[6].strip()

        if (
            len(date) != 10
            or len(transaction_code) != 1
            # or not price.startswith("$")
            # or price == "$0"
        ):
            validation_errors_count += 1
            continue

        row_data = {
            "date": date,
            "transaction_code": transaction_code,
            "price": price,
            "shares_amount": shares_amount,
            "url": url,
            "name": name,
            "total_amount": total_amount,
            "relationship": relationship,
            "ad": ad,
        }

        result.append(row_data)

    logger.debug(f"Errors: {validation_errors_count}/{len(rows)}")

    return result


def write_excel(data: dict, worksheet, row: int) -> None:
    worksheet.write(row, 0, data["name"])
    worksheet.write(row, 1, data["relationship"])
    worksheet.write(row, 2, data["date"])
    worksheet.write(row, 3, data["transaction_code"])
    worksheet.write(row, 4, data["price"])
    worksheet.write(row, 5, data["shares_amount"])
    if isinstance(data["price"], float) and isinstance(data["shares_amount"], float):
        worksheet.write(row, 6, data["shares_amount"])
    worksheet.write(row, 7, data["total_amount"])
    worksheet.write(row, 9, data["ad"])
    worksheet.write(row, 10, data["url"])


def parse_company(entity: str, cicks: str):
    workbook, worksheet = create_excel_tamplate(entity)

    company_data = get_company_data(entity, cicks)
    documents_urls = get_company_urls(company_data, entity, cicks)
    row = 1
    errors = 0
    parsed_wrong = 0
    for document_url in documents_urls:
        print(document_url)
        # try:
        if document_data := parse_document(document_url):
            for row_data in document_data:
                write_excel(row_data, worksheet, row)
                row += 1
        else:
            parsed_wrong += 1
        time.sleep(0.2)
        # except Exception as exc:
        #     time.sleep(1)
        #     errors += 1
        #     print(f"Error: {exc}")
        #     continue

    logger.info(f"Success portion: {(row - 1)/len(documents_urls) * 100}%")
    logger.info(f"Errors portion: {errors/len(documents_urls) * 100}%")
    logger.info(f"Parsed wrong portion: {parsed_wrong/len(documents_urls) * 100}%")

    workbook.close()


def parse_companies(file_name="companies.xlsx"):
    file = openpyxl.load_workbook(file_name)
    wsheet = file.active
    companies = []
    for row in wsheet.iter_rows(max_row=30):
        for cell in row:
            if not cell.value:
                break
            companies.append(cell.value)

    logger.debug(f"Companies found: {companies}")
    for company in companies:
        cik = get_cik(company)
        parse_company(company, cik)


if __name__ == "__main__":
    parse_company("Walmart Inc. (WMT) (CIK 0000104169)", "0000104169")
    # parse_companies()
