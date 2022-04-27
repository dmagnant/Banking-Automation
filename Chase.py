from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, ElementClickInterceptedException
import time
from datetime import datetime
import os
from Functions import setDirectory, chromeDriverAsUser, getUsername, getPassword, openGnuCashBook, showMessage, getGnuCashBalance, updateSpreadsheet, importGnuTransaction

directory = setDirectory()
driver = chromeDriverAsUser(directory)
driver.implicitly_wait(5)
driver.get("https://www.chase.com/")
time.sleep(2)
# login
showMessage("Login Manually", 'login manually \n' 'Then click OK \n')

# # EXPORT TRANSACTIONS
# click on Credit Card (necessary when Checking account also exists)
driver.get("https://secure07a.chase.com/web/auth/dashboard#/dashboard/overviewAccounts/overview/multiProduct;flyout=accountSummary,818208017,CARD,BAC")
# click on Activity since last statement
time.sleep(3)
driver.find_element(By.XPATH, "//*[@id='header-transactionTypeOptions']/span[1]").click()
# choose last statement
driver.find_element(By.ID, "item-0-STMT_CYCLE_1").click()
time.sleep(1)
# # Capture Statement balance
chase = driver.find_element(By.XPATH, "/html/body/div[2]/div/div[23]/div/div[2]/div[1]/div/div/div/div[1]/div/div[1]/div/div[3]/div/div[1]/ul/li[2]/div[1]/div/span[2]").text.strip('$')
time.sleep(1)
# click Download
driver.find_element(By.XPATH, "//*[@id='downloadActivityIcon']").click()
time.sleep(2)
# click Download
driver.find_element(By.ID, "download").click()
# get current date
today = datetime.today()
year = today.year
month = today.month

# # IMPORT TRANSACTIONS
day = today.strftime('%d')
monthto = today.strftime('%m')
yearto = str(year)
monthfrom = "12"         if month == 1 else "{:02d}".format(month - 1)
yearfrom = str(year - 1) if month == 1 else yearto
fromdate = yearfrom + monthfrom + "07_"
todate = yearto + monthto + "06_"
currentdate = yearto + monthto + day
transactions_csv = r'C:\Users\dmagn\Downloads\Chase2715_Activity' + fromdate + todate + currentdate + '.csv'
time.sleep(2)
# Set Gnucash Book
mybook = openGnuCashBook(directory, 'Finance', False, False)

review_trans = importGnuTransaction('Chase', transactions_csv, mybook, driver, directory)
# Back to Accounts page
driver.find_element(By.ID, "backToAccounts").click()
# # REDEEM REWARDS
# open redeem for cash back page
driver.get("https://ultimaterewardspoints.chase.com/cash-back?lang=en")
try:
    # Deposit into a Bank Account
    driver.find_element(By.XPATH, "/html/body/the-app/main/ng-component/main/div/section[2]/div[2]/form/div[6]/ul/li[2]/label").click()
    # Click Continue
    driver.find_element(By.XPATH, "/html/body/the-app/main/ng-component/main/div/section[2]/div[2]/form/div[7]/button").click()
    # Click Confirm & Submit
    driver.find_element(By.ID, "cash_back_button_submit").click()
except NoSuchElementException:
    exception = "caught"
except ElementClickInterceptedException:
    exception = "caught"

chase_gnu = getGnuCashBalance(mybook, 'Chase')
# switch worksheets if running in December (to next year's worksheet)
if month == 12:
    year = year + 1
chase_neg = float(chase) * -1
updateSpreadsheet(directory, 'Checking Balance', year, 'Chase', month, chase_neg)
updateSpreadsheet(directory, 'Checking Balance', year, 'Chase', month, chase_neg, True)

# Display Checking Balance spreadsheet
driver.execute_script("window.open('https://docs.google.com/spreadsheets/d/1684fQ-gW5A0uOf7s45p9tC4GiEE5s5_fjO5E7dgVI1s/edit#gid=1688093622');")
# Open GnuCash if there are transactions to review
if review_trans:
    os.startfile(directory + r"\Finances\Personal Finances\Finance.gnucash")
# Display Balance
showMessage("Balances + Review", f'Chase Balance: {chase} \n' f'GnuCash Chase Balance: {chase_gnu} \n \n' f'Review transactions:\n{review_trans}')
driver.quit()
