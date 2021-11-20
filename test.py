import time
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException
import gspread
from datetime import datetime, timedelta
from decimal import Decimal

from Ally import runAlly
from Functions import setDirectory, chromeDriverAsUser, closeExpressVPN, getPassword, updateSpreadsheet, getCell

directory = setDirectory()
type = "Asset Allocation"
today = datetime.today()
year = today.year
month = 1
value = float(999.99)
account = 'HE_HSA'

# updateSpreadsheet(directory, type, year, account, month, value)

# amex = '$100.25'
# # convert balance from currency (string) to negative amount
# amex_neg = float(amex.strip('$')) * -1

# print(amex_neg)