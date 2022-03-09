import pandas as pd
from constants import *


def read_excel(file_path, **kwargs):
    result = pd.read_excel(file_path, **kwargs)
    return result


def setup_excel_writer():
    writer = pd.ExcelWriter(OUTPUT_EXCEL_FILE_PATH, engine='xlsxwriter')
    pd.DataFrame().to_excel(writer, sheet_name=OUTPUT_EXCEL_FILE_SHEET)
    workbook = writer.book
    worksheet = writer.sheets[OUTPUT_EXCEL_FILE_SHEET]
    return writer, workbook, worksheet


def write_company_info(workbook, worksheet, row, company_name, company_inn, company_mother_name):
    company_info_format = workbook.add_format({'bold': True, 'font_size': 16, 'bg_color': COMPANY_BACKGROUND_COLOR})

    worksheet.write(row, 0, company_name, company_info_format)
    worksheet.write(row, 1, company_inn, company_info_format)
    worksheet.write(row, 2, company_mother_name, company_info_format)


def write_audit_title(workbook, worksheet, row, title_width_in_columns, title):
    audit_title_format = workbook.add_format({'bold': True, 'font_size': 14, 'bg_color': AUDIT_TITLE_BACKGROUND_COLOR})
    worksheet.merge_range(row, 0, row, title_width_in_columns, title, audit_title_format)


def write_audit_table(table, writer, **kwargs):
    table.to_excel(writer, **kwargs)
