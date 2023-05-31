import requests
import xlsxwriter
from loguru import logger
from typing import Tuple, Any, Optional


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


def remove_brakets(string: str) -> str:
    start = string.find("(")
    if start > 0:
        string = string[:start]

    return string


def create_excel_tamplate(entity: str) -> Tuple[xlsxwriter.Workbook, Any]:
    company_name = get_company_name(entity)
    workbook = xlsxwriter.Workbook(f"{company_name}.xlsx")
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
        "AD",
        "Url",
    ]
    for col, column in enumerate(columns):
        worksheet.write(0, col, column)

    return (workbook, worksheet)


def get_company_data(entity: str, cicks: str, page: Optional[int] = None) -> dict:
    url = "https://efts.sec.gov/LATEST/search-index"
    params = {
        "category": "custom",
        "ciks": cicks,
        "entityName": entity,
        "forms": 4,
        "startdt": "2011-01-01",
        "enddt": "2022-12-31",
        "dateRange": "custom",
    }
    if page:
        params["page"] = page
        params["from"] = 100 * (page - 1)

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


def get_company_urls(company_data: dict, entity: str, cicks: str) -> list:
    total_documents = int(company_data["hits"]["total"]["value"])
    if total_documents % 100 == 0:
        num_pages = total_documents // 100
    else:
        num_pages = total_documents // 100 + 1

    documents_urls = []

    for page in range(1, num_pages + 1):
        page_company_data = get_company_data(entity, cicks, page)
        documents_urls_page = [
            get_document_url(document) for document in page_company_data["hits"]["hits"]
        ]
        documents_urls += documents_urls_page

    return documents_urls


def get_cik(name: str) -> str:
    return name[-11:-1]


def get_company_name(name: str) -> str:
    name = name.split()[0]
    name = name.replace("/", "")
    return name
