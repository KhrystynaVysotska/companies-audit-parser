CHROME_EXE_LOCATION = r'C:\Users\ПК\AppData\Local\Google\Chrome\Application\chrome.exe'
CHROME_DRIVE_EXE_LOCATION = r'D:\StandForUkraine\SearchesAnalytics\chromedriver.exe'

AUDIT_PAGE_URL = "https://www.audit-it.ru/buh_otchet/"
AUDIT_PAGE_BASE_URL = "https://www.audit-it.ru"

# INPUT EXCEL FILE CONFIGS
INPUT_EXCEL_FILE_PATH = 'input/companies.xlsx'
INPUT_EXCEL_FILE_SHEET = 'Parsing'

# INPUT EXCEL FILE STRUCTURE CONFIGS
INN_COLUMN_NAME = 'ИНН'
COMPANY_COLUMN_NAME = 'Company'
MOTHER_COMPANY_COLUMN_NAME = 'Mother'

# OUTPUT EXCEL FILE CONFIGS
OUTPUT_EXCEL_FILE_PATH = 'output/companies_audit.xlsx'
OUTPUT_EXCEL_FILE_SHEET = 'Audit'

# ANY NON-NEGATIVE INTEGER OR 'all' TO GET ALL DATA
NUMBER_OF_ITEMS_TO_READ = 'all'

# COLORS
COMPANY_BACKGROUND_COLOR = "#a4c2f4"
AUDIT_TITLE_BACKGROUND_COLOR = "#e8f1ff"
