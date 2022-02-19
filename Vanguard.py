from selenium.webdriver.common.by import By
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
driver.find_element(By.ID, "username").send_keys(getUsername(directory, 'Vanguard'))
time.sleep(1)
# Enter password
driver.find_element(By.ID, "pword").send_keys(getPassword(directory, 'Vanguard'))
time.sleep(1)
# click Log in
driver.find_element(By.XPATH, "//*[@id='vui-button-1']/button/div").click()
# handle security code
try:
    driver.find_element(By.ID, 'vui-radio-1-input-label').click()
    showMessage('Security Code', "Enter Security code, then click OK")
    driver.find_element(By.XPATH, "//*[@id='security-code-submit-btn']/button/div").click()
except NoSuchElementException:
    exception = "caught"
# navigate to asset details page (click view all assets)
driver.get('https://ownyourfuture.vanguard.com/main/dashboard/assets-details')
time.sleep(2)
# move cursor to middle window
pyautogui.moveTo(500, 500)
#scroll down
pyautogui.scroll(-1000)
# Get Total Account Balance
vanguard = driver.find_element(By.XPATH, "/html/body/div[3]/div/app-dashboard-root/app-assets-details/app-balance-details/div/div[3]/div[3]/div/app-details-card/div/div/div[1]/div[3]/h4").text.strip('$').replace(',', '')
# Get Interest YTD
interest_ytd = driver.find_element(By.XPATH, "/html/body/div[3]/div/app-dashboard-root/app-assets-details/app-balance-details/div/div[3]/div[4]/div/app-details-card/div/div/div[1]/div[3]/h4").text.strip('$').replace(',', '')

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
                    if str(tr.post_date.strftime('%Y')) == str(lastmonth[0].year)
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
    writeGnuTransaction(mybook, "Contribution + Interest", lastmonth[1], [-interest, -emp_contribution, account_change], from_account)   
book.close()
vanguard_gnu = getGnuCashBalance(mybook, 'VanguardPension')
updateSpreadsheet(directory, 'Asset Allocation', year, 'VanguardPension', month, float(vanguard))
# Start Gnu cash
os.startfile(directory + r"\Finances\Personal Finances\Finance.gnucash")
# Display Asset Allocation spreadsheet
driver.execute_script("window.open('https://docs.google.com/spreadsheets/d/1sWJuxtYI-fJ6bUHBWHZTQwcggd30RcOSTMlqIzd1BBo/edit#gid=2058576150');")
# display Balances
showMessage("Balances",f'Pension Balance: {vanguard} \n'f'GnuCash Pension Balance: {vanguard_gnu} \n'f'Interest earned: {interest} \n'f'Total monthly contributions: {emp_contribution} \n')
driver.quit()
