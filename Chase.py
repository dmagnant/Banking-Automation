from selenium.common.exceptions import NoSuchElementException, ElementClickInterceptedException
from decimal import Decimal
import time
from datetime import datetime
import csv
from piecash import Transaction, Split
import os
from Functions import setDirectory, chromeDriverAsUser, getUsername, getPassword, openGnuCashBook, showMessage, getGnuCashBalance, updateSpreadsheet, setToAccount

directory = setDirectory()
driver = chromeDriverAsUser(directory)
driver.implicitly_wait(5)
driver.get("https://www.chase.com/")
driver.maximize_window()
time.sleep(2)
# login
showMessage("Login Manually", 'login manually \n' 'Then click OK \n')

# # EXPORT TRANSACTIONS
# click on Credit Card (necessary when Checking account also exists)
driver.get("https://secure07a.chase.com/web/auth/dashboard#/dashboard/overviewAccounts/overview/multiProduct;flyout=accountSummary,818208017,CARD,BAC")
# click on Activity since last statement
time.sleep(3)
driver.find_element_by_xpath("//*[@id='header-transactionTypeOptions']/span[1]").click()
# choose last statement
driver.find_element_by_id("item-0-STMT_CYCLE_1").click()
time.sleep(1)
# # Capture Statement balance
chase = driver.find_element_by_xpath("/html/body/div[2]/div/div[23]/div/div[2]/div[1]/div/div/div/div[4]/div/div[4]/div/div/div[3]/div[1]/dl/dd").text
time.sleep(1)
# click Download
driver.find_element_by_xpath("//*[@id='downloadActivityIcon']").click()
time.sleep(2)
# click Download
driver.find_element_by_id("download").click()
# get current date
today = datetime.today()
year = today.year
month = today.month

# # IMPORT TRANSACTIONS
review_trans = ""
day = today.strftime('%d')
monthto = today.strftime('%m')
if month == 1:
    monthfrom = "12"
    yearto = str(year)
    yearfrom = str(year - 1)
else:
    monthfrom = "{:02d}".format(month - 1)
    yearto = str(year)
    yearfrom = yearto

fromdate = yearfrom + monthfrom + "07_"
todate = yearto + monthto + "06_"
currentdate = yearto + monthto + day
filename = r'C:\Users\dmagn\Downloads\Chase2715_Activity' + fromdate + todate + currentdate + '.csv'
time.sleep(2)
# Set Gnucash Book
mybook = openGnuCashBook(directory, 'Finance', False, False)
# open CSV file at the given path
with open(filename) as csv_file:
    csv_reader = csv.reader(csv_file, delimiter=',')
    line_count = 0
    for row in csv_reader:
        # skip header line
        if line_count == 0:
            line_count += 1
        else:
            # Skip payment (already captured in Checking Balance script
            if "AUTOMATIC PAYMENT" in row[2]:
                continue
            else:
                to_account = setToAccount('Chase', row)
                if to_account == "Expenses:Other":
                    review_trans = review_trans + row[0] + ", " + row[1] + ", " + "\n"
            amount = Decimal(row[5])
            from_account = "Liabilities:Credit Cards:Chase Freedom"
            postdate = datetime.strptime(row[1], '%m/%d/%Y')
            with mybook as book:
                USD = mybook.currencies(mnemonic="USD")
                # create transaction with core objects in one step
                trans = Transaction(post_date=postdate.date(),
                                    currency=USD,
                                    description=row[2],
                                    splits=[
                                         Split(value=-amount, memo="scripted", account=mybook.accounts(fullname=to_account)),
                                         Split(value=amount, memo="scripted", account=mybook.accounts(fullname=from_account)),
                                     ])
                book.save()
                book.flush()
book.close()
# Back to Accounts page
driver.find_element_by_id("backToAccounts").click()
# # REDEEM REWARDS
# open redeem for cash back page
driver.get("https://ultimaterewardspoints.chase.com/cash-back?lang=en")
try:
    # Deposit into a Bank Account
    driver.find_element_by_xpath("/html/body/the-app/main/ng-component/main/div/section[2]/div[2]/form/div[6]/ul/li[2]/label").click()
    # Click Continue
    driver.find_element_by_xpath("/html/body/the-app/main/ng-component/main/div/section[2]/div[2]/form/div[7]/button").click()
    # Click Confirm & Submit
    driver.find_element_by_id("cash_back_button_submit").click()
except NoSuchElementException:
    exception = "caught"
except ElementClickInterceptedException:
    exception = "caught"

chase_gnu = getGnuCashBalance(mybook, 'Chase')
# switch worksheets if running in December (to next year's worksheet)
if month == 12:
    year = year + 1
chase_neg = float(chase.strip('$')) * -1
updateSpreadsheet(directory, 'Checking Balance', year, 'Chase', month, chase_neg)
updateSpreadsheet(directory, 'Checking Balance', year, 'Chase', month, chase_neg, True)

# Display Checking Balance spreadsheet
driver.execute_script("window.open('https://docs.google.com/spreadsheets/d/1684fQ-gW5A0uOf7s45p9tC4GiEE5s5_fjO5E7dgVI1s/edit#gid=914927265');")
# Open GnuCash if there are transactions to review
if review_trans:
    os.startfile(directory + r"\Finances\Personal Finances\Finance.gnucash")
# Display Balance
showMessage("Balances + Review", f'Chase Balance: {chase} \n' f'GnuCash Chase Balance: {chase_gnu} \n \n' f'Review transactions:\n{review_trans}')
driver.quit()