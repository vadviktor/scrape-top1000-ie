import os
import re
from datetime import date

import xlsxwriter
from playwright.sync_api import sync_playwright

BASE_URL = "https://www.top1000.ie"


def text_by(page, selector: str) -> str:
    l = page.locator(selector)
    if l.count() == 0:
        return ""

    return l.inner_text().strip()


workbook = xlsxwriter.Workbook("top1kie.xlsx")
worksheet = workbook.add_worksheet(date.today().strftime("%Y-%m-%d"))
worksheet.write(0, 0, "rank")
worksheet.write(0, 1, "name")
worksheet.write(0, 2, "description")
worksheet.write(0, 3, "employees")
worksheet.write(0, 4, "turnover")
worksheet.write(0, 5, "contact name")
worksheet.write(0, 6, "contact position")

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)
    context = browser.new_context()
    main_page = context.new_page()
    main_page.goto(f"file://{os.path.abspath(os.curdir)}/companies_list.html")
    companies_links = main_page.locator("#companies div.companylisting > a")

    for i in range(companies_links.count()):
        company_page = context.new_page()
        link = companies_links.nth(i).get_attribute("href")
        company_page.goto(f"{BASE_URL}{link}")

        rank = re.sub(
            r"[a-z]+", "", text_by(company_page, "#content div.companyInfo span.rank")
        )
        worksheet.write(i + 1, 0, rank)

        name = text_by(company_page, "#content div.companyDetails h1")
        worksheet.write(i + 1, 1, name)

        description = text_by(
            company_page, "#content div.companyDetails div.description"
        )
        worksheet.write(i + 1, 2, description)

        elem = company_page.locator("span:right-of(label:text('Employees:'))").first
        if elem.count() == 1:
            worksheet.write(i + 1, 3, elem.inner_text().strip())

        elem = company_page.locator("span:right-of(label:text('Turnover:'))").first
        if elem.count() == 1:
            worksheet.write(i + 1, 4, elem.inner_text().strip())

        contact_name = text_by(
            company_page, "#content div.people > ul > li > span.name"
        )
        worksheet.write(i + 1, 5, contact_name)

        contact_position = text_by(
            company_page, "#content div.people > ul > li > span.position"
        )
        worksheet.write(i + 1, 6, contact_position)

        company_page.close()

    browser.close()

workbook.close()
