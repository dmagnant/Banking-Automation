from datetime import datetime
import time
from typing import KeysView
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
import os
from Functions import setDirectory, chromeDriverAsUser, getUsername, getPassword, openGnuCashBook, showMessage, getGnuCashBalance, updateSpreadsheet, importGnuTransaction

directory = setDirectory()
driver = chromeDriverAsUser(directory)
driver.implicitly_wait(3)
driver.get("https://www.barclaycardus.com/servicing/home?secureLogin=")

# Login
driver.find_element(By.ID, "username").send_keys(getUsername(directory, 'Barclay Card'))
driver.find_element(By.ID, "password").send_keys(getPassword(directory, 'Barclay Card'))
driver.find_element(By.ID, "loginButton").click()
# handle security questions
try:
    question1 = driver.find_element(By.ID, "question1TxtbxLabel").text
    question2 = driver.find_element(By.ID, "question2TxtbxLabel").text
    q1 = "In what year was your mother born?"
    a1 = os.environ.get('MotherBirthYear')
    q2 = "What is the name of the first company you worked for?"
    a2 = os.environ.get('FirstEmployer')
    if question1 == q1:
        driver.find_element(By.ID, "rsaAns1").send_keys(a1)
        driver.find_element(By.ID, "rsaAns2").send_keys(a2)
    else:
        driver.find_element(By.ID, "rsaAns1").send_keys(a2)
        driver.find_element(By.ID, "rsaAns2").send_keys(a1)
    driver.find_element(By.ID, "rsaChallengeFormSubmitButton").click()
except NoSuchElementException:
    exception = "Caught"
# handle Confirm Your Identity
try:
    driver.find_element(By.XPATH, "/html/body/section[2]/div[4]/div/div/div[2]/form/div/div[1]/div[2]/div/div/div[1]/div/div/div[2]/div[1]/label/span").click()
    driver.find_element(By.XPATH, "//button[@type='submit']").click()
    # User pop-up to enter SecurPass Code
    showMessage("Get Code From Phone", "Enter in box, then click OK")
    driver.find_element(By.XPATH, "//button[@type='submit']").click()
except NoSuchElementException:
    exception = "Caught"
# handle Pop-up
try:
    driver.find_element(By.XPATH, "/html/body/div[4]/div/button/span").click()
except NoSuchElementException:
    exception = "Caught"
# # Capture Statement balance
barclays = driver.find_element(By.XPATH, "/html/body/section[2]/div[4]/div[2]/div[1]/section[2]/div/div/div[2]/div/div[2]/div[2]/div[2]/div[2]").text.strip('-').strip('$')
# Capture Rewards balance
rewards_balance = driver.find_element(By.XPATH, "//*[@id='rewardsTile']/div[2]/div/div[2]/div[1]/div").text.strip('$')
# # EXPORT TRANSACTIONS
# click on Activity & Statements
driver.find_element(By.XPATH, "/html/body/section[2]/div[1]/nav/div/ul/li[3]/a").click()
# Click on Transactions
driver.find_element(By.XPATH, "/html/body/section[2]/div[1]/nav/div/ul/li[3]/ul/li/div/div[2]/ul/li[1]/a").click()
# Click on Download
driver.find_element(By.XPATH, "/html/body/section[2]/div[4]/div/div/div[3]/div[1]/div/div[2]/span/div/button/span[1]").click()
# get current date
today = datetime.today()
year = today.year
month = today.month
monthto = str(month)
if month == 1:
    monthfrom = "12"
    yrto = str(year - 2000)
    yearto = str(year)
    yrfrom = str(year - 2001)
    yearfrom = str(year - 1)
else:
    monthfrom = str(month - 1)
    yrto = str(year - 2000)
    yearto = str(year)
    yrfrom = yrto
    yearfrom = yearto
fromdate = monthfrom + "/11/" + yrfrom
todate = monthto + "/10/" + yrto
transactions_csv = r"C:\Users\dmagn\Downloads\CreditCard_" + yearfrom + monthfrom + "11_" + yearto + monthto + "10.csv"
# enter date_range
driver.find_element(By.ID, "downloadFromDate_input").send_keys(fromdate)
driver.find_element(By.ID, "downloadToDate_input").send_keys(todate)
# click Download
driver.find_element(By.XPATH, "/html/body/div[3]/div[2]/div/div/div[2]/div/form/div[3]/div/button").click()
# # IMPORT TRANSACTIONS
time.sleep(2)
# Set Gnucash Book
mybook = openGnuCashBook(directory, 'Finance', False, False)
review_trans = importGnuTransaction('Barclays', transactions_csv, mybook, driver, directory, 5)
barclays_gnu = getGnuCashBalance(mybook, 'Barclays')
if float(rewards_balance) > 50:
    # # REDEEM REWARDS
    # click on Rewards & Benefits
    driver.find_element(By.XPATH, "/html/body/section[2]/div[1]/nav/div/ul/li[4]/a").click()
    # click on Redeem my cash rewards
    driver.find_element(By.XPATH, "//*[@id='rewards-benefits-container']/div[1]/ul/li[3]/a").click()
    # click on Direct Deposit or Statement Credit
    driver.find_element(By.XPATH, "/html/body/div[1]/div[2]/div[2]/div[1]/div/div/div[2]/ul/li[1]/div/a").click()
    # click on Continue
    driver.find_element(By.ID, "redeem-continue").click()
    # Click "select an option" (for method of rewards)
    driver.find_element(By.XPATH, "//*[@id='mor_dropDown0']").click()
    driver.find_element(By.XPATH, "//*[@id='mor_dropDown0']").send_keys(KeysView.DOWN)
    driver.find_element(By.XPATH, "//*[@id='mor_dropDown0']").send_keys(KeysView.ENTER)
    time.sleep(1)
    # Click Continue
    driver.find_element(By.XPATH, "//*[@id='achModal-continue']").click()
    time.sleep(1)
    # click on Redeem Now
    driver.find_element(By.XPATH, "/html/body/section[2]/div[4]/div[2]/cashback/div/div[2]/div/ui-view/redeem/div/review/div/div/div/div/div[2]/form/div[3]/div/div[1]/button").click()

# switch worksheets if running in December (to next year's worksheet)
if month == 12:
    year = year + 1
barclays_neg = float(barclays) * -1
updateSpreadsheet(directory, 'Checking Balance', year, 'Barclays', month, barclays_neg, 'Barclays CC')
updateSpreadsheet(directory, 'Checking Balance', year, 'Barclays', month, barclays_neg, 'Barclays CC', True)
# Display Checking Balance spreadsheet
driver.execute_script("window.open('https://docs.google.com/spreadsheets/d/1684fQ-gW5A0uOf7s45p9tC4GiEE5s5_fjO5E7dgVI1s/edit#gid=1688093622');")
# Open GnuCash if there are transactions to review
if review_trans:
    os.startfile(directory + r"\Finances\Personal Finances\Finance.gnucash")
# Display Balance
showMessage("Balances + Review", f'Barclays balance: {barclays} \n' f'GnuCash Barclays balance: {barclays_gnu} \n \n' f'Review transactions:\n{review_trans}')
driver.quit()
