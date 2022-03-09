from audit_parser import get_audit_for, setup_driver
from constants import *
from excel_utils import setup_excel_writer, read_excel, write_company_info, write_audit_title, write_audit_table

if __name__ == "__main__":
    writer, workbook, worksheet = setup_excel_writer()
    companies = read_excel(INPUT_EXCEL_FILE_PATH, sheet_name=INPUT_EXCEL_FILE_SHEET, dtype={INN_COLUMN_NAME: str})

    if NUMBER_OF_ITEMS_TO_READ == 'all':
        companies_inns = companies[INN_COLUMN_NAME]
        companies_names = companies[COMPANY_COLUMN_NAME]
        companies_mother_names = companies[MOTHER_COMPANY_COLUMN_NAME]
    else:
        companies_inns = companies[INN_COLUMN_NAME][:NUMBER_OF_ITEMS_TO_READ]
        companies_names = companies[COMPANY_COLUMN_NAME][:NUMBER_OF_ITEMS_TO_READ]
        companies_mother_names = companies[MOTHER_COMPANY_COLUMN_NAME][:NUMBER_OF_ITEMS_TO_READ]

    companies_counter = 1
    rows_counter = 0

    browser = setup_driver()
    for row in zip(companies_inns, companies_names, companies_mother_names):
        inn, company, mother = row[0], row[1], row[2]

        print(companies_counter, company)

        # check if inn is nan (float type), otherwise it has str type
        if type(inn) == float:
            write_company_info(workbook, worksheet, rows_counter, company, "", mother)
            rows_counter += 1
        else:
            write_company_info(workbook, worksheet, rows_counter, company, inn, mother)
            rows_counter += 1

            audit = get_audit_for(inn, browser)

            for table_name, table in audit:
                if len(table.columns):
                    write_audit_title(workbook, worksheet, rows_counter, len(table.columns) - 1, table_name)
                    rows_counter += 1

                    write_audit_table(table, writer, startrow=rows_counter, sheet_name=OUTPUT_EXCEL_FILE_SHEET,
                                      index=False)
                    rows_counter += len(table) + 2

        rows_counter += 1
        companies_counter += 1

    browser.quit()
    writer.save()
