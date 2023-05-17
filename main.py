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


def parse_document(url: str) -> list:
    res = requests.get(url, headers=HEADERS)
    if res.status_code != 200:
        logger.warning(f"Could not get document with url {url}")
        return
    document = BeautifulSoup(res.content, "html.parser")
    table_1: Tag = document.find(
        "b",
        string="Table I - Non-Derivative Securities Acquired, Disposed of, or Beneficially Owned",
    ).parent.parent.parent.parent
    table_1_data = table_1.select_one("tbody").select_one("tr")
    table_1_data = [
        column_raw.text
        for column_raw in table_1_data
        if column_raw.text not in ["\n", ""]
    ]
    date = table_1_data[1]
    transaction_code = table_1_data[2]
    price = table_1_data[5]
    shares_amount = table_1_data[3]

    #name = document.select_one("span.FormData").text

    return {
        "date": date,
        "transaction_code": transaction_code,
        "price": price,
        "shares_amount": shares_amount,
        #"name": name
    }


def write_excel(data: dict, worksheet, row: int) -> None:
    # worksheet.write(row, 0, data["name"])
    if len(data["date"]) != 10 or len(data["transaction_code"]):
        return

    worksheet.write(row, 2, data["date"])
    worksheet.write(row, 3, data["transaction_code"])
    worksheet.write(row, 4, data["price"])
    worksheet.write(row, 5, data["shares_amount"])
    # worksheet.write(row, 7, data["price"])


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
]
for col, column in enumerate(columns):
    worksheet.write(0, col, column)

json = get_entity_data("Apple Inc. (AAPL) (CIK 0000320193)", "0000320193")
documents_urls = [get_document_url(document) for document in json["hits"]["hits"]]
for row, document_url in enumerate(documents_urls):
    print(document_url)
    try:
        document_dict = parse_document(document_url)
        #pprint.pp(document_dict)
        write_excel(document_dict, worksheet, row + 1)
    except:
        continue

workbook.close()
