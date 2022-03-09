import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait

from constants import *


def setup_driver():
    opts = Options()
    opts.binary_location = CHROME_EXE_LOCATION
    chrome_driver = CHROME_DRIVE_EXE_LOCATION
    driver = webdriver.Chrome(chrome_driver, options=opts)
    driver.maximize_window()
    time.sleep(1)
    return driver


def get_capital_changes_data(source_page):
    capital_changes = source_page.find("div", {"id": "form5"})
    if capital_changes:
        # 2020 by default
        year = capital_changes.h3.span.find_all("span")[1].get_text()
        title = capital_changes.h3.label.get_text() + year

        table = capital_changes.table.tbody
        table_rows = table.find_all("tr")

        column_names = table_rows[0].find_all("th")
        column_names = [th.get_text() for th in column_names]

        data = []

        for row in table_rows[1:]:
            if row.has_attr("style"):
                cells = row.find_all("td")
                row_title = cells[0].get_text()
                row_code = cells[1].get_text()
                row_values = [cell.get_text() if cell.get_text() else "-" for cell in cells[2:]]
                data.append((row_title, row_code, *row_values))

        df = pd.DataFrame.from_records(data=data, columns=column_names)
        return title, df
    else:
        return "", pd.DataFrame()


def get_financial_indicators_data(source_page, index):
    table = source_page.find_all("table", class_="tblFin")[index]
    column_names = table.thead.tr.find_all('th')[:2]
    column_names = [th.get_text() for th in column_names]

    data = []

    for row in table.tbody.find_all('tr'):
        cells = row.find_all("td")
        row_title = cells[0].get_text()

        # for 2020 year only
        row_value = cells[1].get_text() if cells[1].get_text() else "-"
        data.append((row_title, row_value))

    df = pd.DataFrame.from_records(data=data, columns=column_names)
    return df


def get_audit_data(source_page, div_id):
    sheet = source_page.find("div", {"id": div_id})
    if sheet:
        title = sheet.h3.get_text() if sheet.h3 else ""
        table = sheet.table.tbody

        table_rows = table.find_all("tr")

        column_names = table_rows[0].find_all("th")[:3]
        column_names = [th.get_text() for th in column_names]

        data = []

        for row in table_rows[1:]:
            if row.has_attr("class") and 'calcRow' in row["class"]:
                cells = row.find_all("td")
                row_title = cells[0].get_text()
                row_code = cells[1].get_text()

                # for 2020 year only
                row_value = cells[2].get_text() if cells[2].get_text() else "-"

                data.append((row_title, row_code, row_value))
            else:
                data.append((row.td.get_text(), "", ""))

        df = pd.DataFrame.from_records(data=data, columns=column_names)
        return title, df
    else:
        return "", pd.DataFrame()


def get_company_table(company_code, browser):
    browser.get(AUDIT_PAGE_URL)

    inn_inputs = browser.find_elements(By.CLASS_NAME, value="buhotchet-search")
    for input in inn_inputs:
        try:
            WebDriverWait(browser, 1).until(EC.element_to_be_clickable(input))
            input.click()
            input.send_keys(company_code)

            search_button = browser.find_element(by=By.XPATH, value="//*[contains(text(),'Найти')]")
            time.sleep(1)
            ActionChains(browser).move_to_element(search_button).click().perform()
            time.sleep(1)
            break
        except Exception as e:
            pass

    page_with_company_table = BeautifulSoup(browser.page_source, features="html.parser")
    company_table = page_with_company_table.find("table", class_="resultsTable")
    return company_table


def get_audit_for(company_code, browser):
    company_table = get_company_table(company_code, browser)

    company_audit = []

    if company_table:
        company_table_values = company_table.tbody.find_all("tr")[1]
        company_audit_url_path = company_table_values.find_all("td")[1]
        company_audit_url_path = company_audit_url_path.a["href"]

        browser.get(AUDIT_PAGE_BASE_URL + company_audit_url_path)
        company_audit_page = BeautifulSoup(browser.page_source, 'html.parser')

        balance_sheet_name, balance_sheet_df = get_audit_data(company_audit_page, "form1")
        company_audit.append((balance_sheet_name, balance_sheet_df))

        balance_financial_indicators_df = get_financial_indicators_data(company_audit_page, 0)
        company_audit.append(("Фінансові показники", balance_financial_indicators_df))

        profits_and_expenses_sheet_name, profits_and_expenses_sheet_df = get_audit_data(company_audit_page, "form2")
        company_audit.append((profits_and_expenses_sheet_name, profits_and_expenses_sheet_df))

        profits_and_expenses_financial_indicators_df = get_financial_indicators_data(company_audit_page, 1)
        company_audit.append(("Фінансові показники", profits_and_expenses_financial_indicators_df))

        money_movement_sheet_name, money_movement_sheet_df = get_audit_data(company_audit_page, "form4")
        company_audit.append((money_movement_sheet_name, money_movement_sheet_df))

        capital_changes_name, capital_changes_df = get_capital_changes_data(company_audit_page)
        company_audit.append((capital_changes_name, capital_changes_df))

        capital_changes_other_name, capital_changes_other_df = get_audit_data(company_audit_page, "form3")
        company_audit.append((capital_changes_other_name, capital_changes_other_df))

    return company_audit
