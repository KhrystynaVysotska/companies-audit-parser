# Audit parser
![image](https://user-images.githubusercontent.com/56559854/157432409-46d83507-97bf-4891-803e-bd6ca9633cfa.png)

### Features
- Retrieves 2020 audit data of companies that have their businesses in Russia.
- Scrapes data from [Russia public audit page](https://www.audit-it.ru/buh_otchet/) 
- Uses company's INN to get audit data
- Creates excel file with gathered data for all provided companies


### Configuration
| Constant                      | Description                                                                                                 |
| ----------------------------- | ----------------------------------------------------------------------------------------------------------  |
| CHROME_EXE_LOCATION           | Path to `chrome.exe`. Must be changed to local path before script execution                                 |
| CHROME_DRIVE_EXE_LOCATION     | Path to `chromedriver.exe`. Must be changed to local path before script execution                           |
| AUDIT_PAGE_URL                | URL of page with audit data. Shouldn't be changed. Defaults to `https://www.audit-it.ru/buh_otchet/`        |
| AUDIT_PAGE_BASE_URL           | Base URL of page with audit data. Shouldn't be changed. Defaults to `https://www.audit-it.ru`               |
| INPUT_EXCEL_FILE_PATH         | Path to input excel file. Defaults to `input/companies.xlsx`                                                |
| INPUT_EXCEL_FILE_SHEET        | Sheet that will be created in output excel file. Default to `Parsing`                                       |
| INN_COLUMN_NAME               | Name of INN column in input excel file. Defaults to `ИНН`                                                   |
| COMPANY_COLUMN_NAME           | Name of Company column in input excel file. Defaults to `Company`                                           |
| MOTHER_COMPANY_COLUMN_NAME    | Name of Mother company name column in input excel file. Defaults to `Mother`                                |
| OUTPUT_EXCEL_FILE_PATH        | Path to output excel file. Defaults to `output/companies_audit.xlsx`                                        |
| OUTPUT_EXCEL_FILE_SHEET       | Sheet that will be created in output excel file. Default to `Audit`                                         |
| NUMBER_OF_ITEMS_TO_READ       | Number of first companies to gather data for. Defaults to `all`. Can be changed to any non-negative integer |
| COMPANY_BACKGROUND_COLOR      | Background color for rows with company info (name, inn, mother company). Defaults to `#a4c2f4`              |
| AUDIT_TITLE_BACKGROUND_COLOR  | Background color for rows with audit table titles. Defaults to `#e8f1ff`                                    |


### How to use
1. Download this repository
2. Install dependencies from requirements.txt
3. Check configs in `constants.py`. Customize them. Remember to change chrome.exe and chromedriver.exe paths.
4. Provide input excel file to `input/companies.xlsx` or custom path (set in INPUT_EXCEL_FILE_PATH). Column names can be set in `constants.py`. Ensure that input excel file has following columns:
    - Company
    - ИНН
    - Mother
5. Run `main.py`
6. The chrome tab will be opened in a new window but you can collapse it and keep doing other things while data is being gathered.
7. Company name and its sequence number are printed to the output so you can check the status of data gathering there.
8. After the data is gathered you can check results in `output/companies_audit.xlsx` (or any custom path set in OUTPUT_EXCEL_FILE_PATH).
