from ahk import AHK
from datetime import datetime, time
import time
from selenium.common.exceptions import NoSuchElementException, ElementNotInteractableException
from decimal import Decimal
import csv
from piecash import Transaction, Split
import os
import pyautogui
from Functions import setDirectory, chromeDriverAsUser, getUsername, getPassword, openGnuCashBook, showMessage, getGnuCashBalance, updateSpreadsheet, setToAccount, importGnuTransaction

directory = setDirectory()
driver = chromeDriverAsUser(directory)
driver.implicitly_wait(5)
driver.get("https://www.discover.com/")
driver.maximize_window()
# login
ahk = AHK()
pyautogui.leftClick(1131, 391)
pyautogui.press('tab')
pyautogui.write(getUsername(directory, 'Discover'))
pyautogui.press('tab')
pyautogui.write(getPassword(directory, 'Discover'))
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.press('enter')
# ahk.mouse_move(x=1131, y=391)
# ahk.click()
# ahk.key_press('Tab')
# ahk.type(entry.username)
# ahk.key_press('Tab')
# ahk.type(entry.password)
# ahk.key_press('Tab')
# ahk.key_press('Tab')
# ahk.key_press('Tab')
# ahk.key_press('Enter')
#handle pop-up
try:
    driver.find_element_by_xpath("/html/body/div[1]/main/div[12]/div/div/div[2]/a").click()
except NoSuchElementException:
    exception = "caught"
except ElementNotInteractableException:
    exception = "caught"
showMessage("Login Check", 'Confirm Login to , (manually if necessary) \n' 'Then click OK \n')
# # Capture Statement balance
discover = driver.find_element_by_xpath("/html/body/div[1]/main/div[5]/div[1]/div[1]/div[1]/p[3]/span[2]").text
# # Export Transactions
# Click on All Activity & Statements
driver.find_element_by_partial_link_text("All Activity & Statements").click()
# Click on "Select Activity or Statement Period"
driver.find_element_by_xpath("/html/body/div[1]/main/div[1]/div[4]/div/div[1]/div/div/div/div/div[1]/a").click()
# Click on Current
driver.find_element_by_partial_link_text("Current").click()
driver.implicitly_wait(3)
# Click on Download
driver.find_element_by_xpath("/html/body/div[1]/main/div[1]/aside/div[2]/div[1]/div/a[2]").click()
# CLick on CSV
driver.find_element_by_id("radio4").click()
# CLick Download
driver.find_element_by_id("submitDownload").click()
# Click Close
driver.find_element_by_xpath("/html/body/div[1]/main/div[5]/div/form/div/div[4]/a[1]").click()
# get current date
today = datetime.today()
year = today.year
month = today.month
# Set Gnucash Book
mybook = openGnuCashBook(directory, 'Finance', False, False)

# # IMPORT TRANSACTIONS
stmtyear = str(year)
stmtmonth = today.strftime('%m')
transactions_csv = r"C:\Users\dmagn\Downloads\Discover-Statement-" + stmtyear + stmtmonth + "12.csv"

review_trans = importGnuTransaction('Discover', transactions_csv, mybook, today)

# Redeem Rewards
# Click Rewards
driver.find_element_by_xpath("/html/body/div[1]/header/div/div/div[3]/div[1]/ul/li[4]/a").click()
# Click "Redeem Cashback Bonus"
driver.find_element_by_xpath("/html/body/div[1]/header/div/div/div[3]/div[1]/ul/li[4]/div/div/div/ul/li[3]/a").click()
# Click "Cash It"
driver.find_element_by_xpath("//*[@id='redemption-module']/li[1]/a").click()
time.sleep(1)
try:
    # Click Electronic Deposit to your bank account
    driver.find_element_by_xpath("//*[@id='electronic-deposit']").click()
    time.sleep(1)
    # Click Redeem All link
    driver.find_element_by_xpath("/html/body/div[1]/div[1]/main/div/div/section/div[2]/div/form/div[2]/fieldset/div[3]/div[2]/span[2]/button").click()
    time.sleep(1)
    # Click Continue
    driver.find_element_by_xpath("/html/body/div[1]/div[1]/main/div/div/section/div[2]/div/form/div[4]/input").click()
    time.sleep(1)
    # Click Submit
    driver.find_element_by_xpath("/html/body/div[1]/div[1]/main/div/div/section/div[2]/div/div/div/div[1]/div/div/div[2]/div/button[1]").click()
except NoSuchElementException:
    exception = "caught"

discover_gnu = getGnuCashBalance(mybook, 'Discover')
# switch worksheets if running in December (to next year's worksheet)
if month == 12:
    year = year + 1
discover_neg = float(discover.strip('$')) * -1
updateSpreadsheet(directory, 'Checking Balance', year, 'Discover', month, discover_neg)
updateSpreadsheet(directory, 'Checking Balance', year, 'Discover', month, discover_neg, True)

# Display Checking Balance spreadsheet
driver.execute_script("window.open('https://docs.google.com/spreadsheets/d/1684fQ-gW5A0uOf7s45p9tC4GiEE5s5_fjO5E7dgVI1s/edit#gid=914927265');")
# Open GnuCash if there are transactions to review
if review_trans:
    os.startfile(directory + r"\Finances\Personal Finances\Finance.gnucash")
# Display Statement Balance
showMessage("Balances + Review", f'Discover Balance: {discover} \n' f'GnuCash Discover Balance: {discover_gnu} \n \n' f'Review transactions:\n{review_trans}')
driver.quit()