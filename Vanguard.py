from selenium.common.exceptions import NoSuchElementException
from decimal import Decimal
from datetime import datetime
import time
import os
import pyautogui
from Functions import setDirectory, chromeDriverAsUser, getUsername, getPassword, openGnuCashBook, showMessage, getGnuCashBalance, updateSpreadsheet, getStartAndEndOfPreviousMonth, writeGnuTransaction 

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
vanguard = driver.find_element_by_xpath("/html/body/div[3]/div/app-dashboard-root/app-assets-details/app-balance-details/div/div[3]/div[3]/div/app-details-card/div/div/div[1]/div[3]/h4").text.strip('$').strip(',')
# Get Interest YTD
interest_ytd = driver.find_element_by_xpath("/html/body/div[3]/div/app-dashboard-root/app-assets-details/app-balance-details/div/div[3]/div[4]/div/app-details-card/div/div/div[1]/div[3]/h4").text.strip('$').strip(',')

#get current date
today = datetime.today()
year = today.year
month = today.month

lastmonth = getStartAndEndOfPreviousMonth(today, month, year)

interest_amount = 0

# Set Gnucash Book
mybook = openGnuCashBook(directory, 'Finance', False, False)

pension = getGnuCashBalance(mybook, 'VanguardPension')
pension_acct = "Assets:Non-Liquid Assets:Pension"
interest_acct = "Income:Investments:Interest"
with mybook as book:
    # # GNUCASH
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
    writeGnuTransaction(mybook, "Contribution + Interest", lastmonth[1].date(), [-interest, -emp_contribution, account_change], from_account)   
    # Transaction(post_date=lastmonth[1].date(),
    #                 currency=mybook.currencies(mnemonic="USD"),
    #                 description="Contribution + Interest",
    #                 splits=[
    #                     Split(value=-interest, memo="scripted",
    #                             account=mybook.accounts(fullname="Income:Investments:Interest")),
    #                     Split(value=-emp_contribution, memo="scripted",
    #                             account=mybook.accounts(fullname="Income:Employer Pension Contributions")),
    #                     Split(value=account_change, memo="scripted",
    #                             account=mybook.accounts(fullname=from_account)),
    #                 ])
    # book.save()
    # book.flush()
book.close()
vanguard_gnu = getGnuCashBalance(mybook, 'VanguardPension')
updateSpreadsheet(directory, 'Asset Allocation', year, 'VanguardPension', month, vanguard)
# Start Gnu cash
os.startfile(directory + r"\Finances\Personal Finances\Finance.gnucash")
# Display Asset Allocation spreadsheet
driver.execute_script("window.open('https://docs.google.com/spreadsheets/d/1sWJuxtYI-fJ6bUHBWHZTQwcggd30RcOSTMlqIzd1BBo/edit#gid=953264104');")
# display Balances
showMessage("Balances",f'Pension Balance: {vanguard} \n'f'GnuCash Pension Balance: {vanguard_gnu} \n'f'Interest earned: {interest} \n'f'Total monthly contributions: {emp_contribution} \n')
driver.quit()