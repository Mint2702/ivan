import time
import pprint
import requests
import xlsxwriter
from bs4 import BeautifulSoup, Tag

from loguru import logger


HEADERS = {
    "accept": "application/json, text/javascript, */*; q=0.01",
    "accept-encoding": "gzip, deflate, br",
    "accept-language": "ru,en;q=0.9",
    "origin": "https://www.sec.gov",
    "referer": "https://www.sec.gov/",
    "sec-ch-ua": '"Chromium";v="104", " Not A;Brand";v="99", "Yandex";v="22"',
    "sec-ch-ua-mobile": "?0",
    "sec-ch-ua-platform": '"Linux"',
    "sec-fetch-mode": "cors",
    "sec-fetch-site": "same-site",
    "user-agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/104.0.5112.114 YaBrowser/22.9.1.1110 (beta) Yowser/2.5 Safari/537.36",
}


def get_entity_data(entity: str, cicks: str) -> str:
    url = "https://efts.sec.gov/LATEST/search-index"
    params = {
        "category": "form-cat2",
        "ciks": cicks,
        "entityName": entity,
        "filter_forms": 4,
        "startdt": "2018-05-05",
        "enddt": "2023-05-05",
    }

    res = requests.get(url, params=params, headers=HEADERS)

    if res.status_code == 200:
        return res.json()
    else:
        logger.warning(
            f"Could not get html for page with entity == {entity}. Got {res.status_code} status code"
        )


def get_document_url(document: dict) -> list:
    first_num = int(document["_source"]["ciks"][1])
    adsh: str = document["_source"]["adsh"].replace("-", "")
    xsl = document["_source"]["xsl"]
    file = document["_id"].split(":")[-1]
    return f"https://www.sec.gov/Archives/edgar/data/{first_num}/{adsh}/{xsl}/{file}"


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
    for row in rows:
        date = row[1].strip()
        transaction_code = row[3].strip()
        shares_amount = row[5].strip()
        price = row[7].strip()
        if price == "(1)":
            continue
        total_amount: str = row[8].strip()
        if total_amount.endswith(")"):
            total_amount = total_amount[:-3]
        name = (
            document.select_one('span:-soup-contains("Reporting Person")')
            .find_next("table")
            .text
        )
        relationship = (
            document.select_one('span:-soup-contains("Relationship")')
            .find_next("table")
            .select_one("tr")
            .select("td")[1]
            .text
        )

        if (
            len(date) != 10
            or len(transaction_code) != 1
            or not price.startswith("$")
            or price == "$0"
        ):
            logger.debug("Validation not passed")
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
        }

        result.append(row_data)

    return result


def write_excel(data: dict, worksheet, row: int) -> None:
    worksheet.write(row, 0, data["name"])
    worksheet.write(row, 1, data["relationship"])
    worksheet.write(row, 2, data["date"])
    worksheet.write(row, 3, data["transaction_code"])
    worksheet.write(row, 4, data["price"])
    worksheet.write(row, 5, data["shares_amount"])
    worksheet.write(row, 7, data["total_amount"])
    worksheet.write(row, 9, data["url"])


workbook = xlsxwriter.Workbook("hello.xlsx")
worksheet = workbook.add_worksheet()
columns = [
    "Insider trading",
    "Relationship",
    "Date",
    "Transaction",
    "Price",
    "Shares (amount)",
    "Value",
    "Shares total",
    "SEC Form 4",
    "Url",
]
for col, column in enumerate(columns):
    worksheet.write(0, col, column)

json = get_entity_data("Apple Inc. (AAPL) (CIK 0000320193)", "0000320193")
documents_urls = [get_document_url(document) for document in json["hits"]["hits"]]
row = 1
errors = 0
parsed_wrong = 0
for document_url in documents_urls:
    print(document_url)
    try:
        if document_data := parse_document(document_url):
            for row_data in document_data:
                write_excel(row_data, worksheet, row)
                row += 1
        else:
            parsed_wrong += 1
        time.sleep(0.02)
    except Exception as exc:
        time.sleep(1)
        errors += 1
        print(f"Error: {exc}")
        continue

logger.info(f"Success portion: {(row - 1)/len(documents_urls) * 100}%")
logger.info(f"Errors portion: {errors/len(documents_urls) * 100}%")
logger.info(f"Parsed wrong portion: {parsed_wrong/len(documents_urls) * 100}%")

workbook.close()
