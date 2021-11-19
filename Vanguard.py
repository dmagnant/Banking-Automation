from selenium.common.exceptions import NoSuchElementException
import gspread
from decimal import Decimal
from datetime import datetime
import time
from piecash import Transaction, Split
import os
import pyautogui
from Functions import setDirectory, chromeDriverAsUser, getUsername, getPassword, openGnuCashBook, showMessage, getGnuCashBalance

directory = setDirectory()
driver = chromeDriverAsUser(directory)
driver.implicitly_wait(5)
driver.get("https://ownyourfuture.vanguard.com/login#/")
driver.maximize_window()
# Enter username
driver.find_element_by_id("username").send_keys(getUsername(directory, 'Vanguard'))
time.sleep(1)
# Enter password
driver.find_element_by_id("password").send_keys(getPassword(directory, 'Vanguard'))
time.sleep(1)
# click SUBMIT
driver.find_element_by_xpath("/html/body/app-root/app-layout/div/div[2]/main/app-login/div/div[1]/div[2]/div/form/vui-button/button").click()
# handle security code
try:
    driver.find_element_by_xpath("//*[@id='LoginForm:DEVICE:0']").click()
    showMessage('Security Code', "Enter Security code, then click OK")
    driver.find_element_by_xpath("//*[@id='LoginForm:ContinueInput']").click()
except NoSuchElementException:
    exception = "caught"
#click View all assets
driver.find_element_by_xpath("/html/body/div[3]/div/app-dashboard-root/app-dashboard/div/div[1]/div/div/div/div/div/div[2]/div[2]/button[2]").click()
time.sleep(2)
# move cursor to middle window
pyautogui.moveTo(500, 500)
#scroll down
pyautogui.scroll(-1000)
# Get Total Account Balance
vanguard = driver.find_element_by_xpath("/html/body/div[3]/div/app-dashboard-root/app-assets-details/app-balance-details/div/div[3]/div[3]/div/app-details-card/div/div/div[1]/div[3]/h4").text.replace('$', '').replace(',', '')
# Get Interest YTD
interest_ytd = driver.find_element_by_xpath("/html/body/div[3]/div/app-dashboard-root/app-assets-details/app-balance-details/div/div[3]/div[4]/div/app-details-card/div/div/div[1]/div[3]/h4").text.replace('$', '').replace(',', '')

#get current date
today = datetime.today()
year = today.strftime('%Y')
month = today.month
month = today.month

if month == 1:
    year = Decimal(year) - 1

# Change postdate to last day of last month
if month == 1:
    postdate = today.replace(month=int(12), day=int(31), year=int(year))
elif month == 2:
    postdate = today.replace(month=int(1), day=int(31))
elif month == 3:
    postdate = today.replace(month=int(2), day=int(28))
elif month == 4:
    postdate = today.replace(month=int(3), day=int(31))
elif month == 5:
    postdate = today.replace(month=int(4), day=int(30))
elif month == 6:
    postdate = today.replace(month=int(5), day=int(31))
elif month == 7:
    postdate = today.replace(month=int(6), day=int(30))
elif month == 8:
    postdate = today.replace(month=int(7), day=int(31))
elif month == 9:
    postdate = today.replace(month=int(8), day=int(31))
elif month == 10:
    postdate = today.replace(month=int(9), day=int(30))
elif month == 11:
    postdate = today.replace(month=int(10), day=int(31))
else:
    postdate = today.replace(month=int(11), day=int(30))

interest_amount = 0

# Set Gnucash Book
mybook = openGnuCashBook(directory, 'Finance', False, False)
pension = getGnuCashBalance(mybook, 'Vanguard')
with mybook as book:
    USD = mybook.currencies(mnemonic="USD")
    # # GNUCASH
    pension_acct = "Assets:Non-Liquid Assets:Pension"
    interest_acct = "Income:Investments:Interest"
    # retrieve transactions from GnuCash
    transactions = [tr for tr in mybook.transactions
                    if str(tr.post_date.strftime('%Y')) == str(year)
                    for spl in tr.splits
                    if spl.account.fullname == pension_acct
                    ]
    for tr in transactions:
        date = str(tr.post_date.strftime('%Y'))
        for spl in tr.splits:
            if spl.account.fullname == interest_acct:
                interest_amount = interest_amount + abs(spl.value)
    interest = Decimal(interest_ytd) - interest_amount
    account_change = Decimal(vanguard) - pension
    emp_contribution = account_change - interest
    from_account = "Assets:Non-Liquid Assets:Pension"
    entry = Transaction(post_date=postdate.date(),
                        currency=USD,
                        description="Contribution + Interest",
                        splits=[
                            Split(value=-interest, memo="scripted",
                                  account=mybook.accounts(fullname="Income:Investments:Interest")),
                            Split(value=-emp_contribution, memo="scripted",
                                  account=mybook.accounts(fullname="Income:Employer Pension Contributions")),
                            Split(value=account_change, memo="scripted",
                                  account=mybook.accounts(fullname=from_account)),
                        ])
    book.save()
    book.flush()
book.close()
vanguard_gnu = getGnuCashBalance(mybook, 'Vanguard')
# add to Asset Allocation spreadsheet
json_creds = directory + r"\Projects\Coding\Python\BankingAutomation\Resources\creds.json"
sheet = gspread.service_account(filename=json_creds).open("Asset Allocation")
year = today.strftime('%Y')
worksheet = sheet.worksheet(str(year))
# update appropriate month's information
if month == 1:
    worksheet.update('B8', vanguard_gnu)
elif month == 2:
    worksheet.update('I8', vanguard_gnu)
elif month == 3:
    worksheet.update('P8', vanguard_gnu)
elif month == 4:
    worksheet.update('B30', vanguard_gnu)
elif month == 5:
    worksheet.update('I30', vanguard_gnu)
elif month == 6:
    worksheet.update('P30', vanguard_gnu)
elif month == 7:
    worksheet.update('B52', vanguard_gnu)
elif month == 8:
    worksheet.update('I52', vanguard_gnu)
elif month == 9:
    worksheet.update('P52', vanguard_gnu)
elif month == 10:
    worksheet.update('B74', vanguard_gnu)
elif month == 11:
    worksheet.update('I74', vanguard_gnu)
else:
    worksheet.update('P74', vanguard_gnu)
# Start Gnu cash
os.startfile(directory + r"\Finances\Personal Finances\Finance.gnucash")
# Display Asset Allocation spreadsheet
driver.execute_script("window.open('https://docs.google.com/spreadsheets/d/1sWJuxtYI-fJ6bUHBWHZTQwcggd30RcOSTMlqIzd1BBo/edit#gid=953264104');")

# display Balances
showMessage("Balances",f'Pension Balance: {vanguard} \n'f'GnuCash Pension Balance: {vanguard_gnu} \n'f'Interest earned: {interest} \n'f'Total monthly contributions: {emp_contribution} \n')
driver.quit()