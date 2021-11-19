import gspread
import pyautogui
from selenium.common.exceptions import NoSuchElementException
from datetime import date, datetime
import time
from decimal import Decimal
import csv
from piecash import Transaction, Split
import os
from ahk import AHK
from Functions import setDirectory, chromeDriverAsUser, getUsername, getPassword, openGnuCashBook, showMessage, getGnuCashBalance

directory = setDirectory()
driver = chromeDriverAsUser(directory)
driver.implicitly_wait(5)
driver.get("https://www.americanexpress.com/")
driver.maximize_window()
# login
driver.find_element_by_xpath("/html/body/div[1]/div/div/header/div[2]/div[1]/div[2]/div/div[6]/ul/li[3]/span/a[1]/span").click()
driver.find_element_by_id("eliloUserID").send_keys(getUsername(directory, 'Amex'))
driver.find_element_by_id("eliloPassword").send_keys(getPassword(directory, 'Amex'))
driver.find_element_by_id("loginSubmit").click()
# handle pop-up
try:
    driver.find_element_by_xpath("/html/body/div[1]/div[5]/div/div/div/div/div/div[2]/div/div/div/div/div[1]/div/a/span/span").click()
except NoSuchElementException:
    exception = "caught"
time.sleep(1)
# # Capture Statement balance
amex = driver.find_element_by_xpath("//*[@id='axp-balance-payment']/div[1]/div[1]/div/div[1]/div/div/span[1]/div").text
# # EXPORT TRANSACTIONS
# click on View Transactions
driver.find_element_by_xpath("//*[@id='axp-balance-payment']/div[2]/div[2]/div/div[1]/div[1]/div/a").click()
try: 
    # click on View Activity (for previous billing period)
    driver.find_element_by_xpath("//*[@id='root']/div[1]/div/div[2]/div/div/div[4]/div/div[3]/div/div/div/div/div/div/div[2]/div/div/div[5]/div/div[2]/div/div[2]/a/span").click()
except NoSuchElementException:
    exception = "caught"
# click on Download
time.sleep(5)
driver.find_element_by_xpath("/html/body/div[1]/div[1]/div/div[2]/div/div/div[4]/div/div[3]/div/div/div/div/div/div/div[2]/div/div/div[2]/div/div/div/div/table/thead/div/tr[1]/td[2]/div/div[2]/div/button").click()
# click on CSV option
driver.find_element_by_xpath("/html/body/div[1]/div[4]/div/div/div/div/div/div[2]/div/div[1]/div/fieldset/div[2]/label").click()
# delete old csv file, if present
try:
    os.remove(r"C:\Users\dmagn\Downloads\activity.csv")
except FileNotFoundError:
    exception = "caught"
# click on Download
driver.find_element_by_xpath("/html/body/div[1]/div[4]/div/div/div/div/div/div[3]/a").click()

# # IMPORT TRANSACTIONS
review_trans = ""
# open CSV file at the given path
time.sleep(3)

# Set Gnucash Book
mybook = openGnuCashBook(directory, 'Finance', False, False)

with open(r'C:\Users\dmagn\Downloads\activity.csv') as csv_file:
    csv_reader = csv.reader(csv_file, delimiter=',')
    line_count = 0
    for row in csv_reader:
        # skip header line
        if line_count == 0:
            line_count += 1
        else:
            if "AUTOPAY PAYMENT" in row[1]:
                continue
            elif "PICK N SAVE" in row[1]:
                to_account = "Expenses:Groceries"
            elif "KOPPA" in row[1]:
                to_account = "Expenses:Groceries"
            elif "KETTLE RANGE MEAT" in row[1]:
                to_account = "Expenses:Groceries"
            elif "BP#" in row[1]:
                to_account = "Expenses:Transportation:Gas (Vehicle)"
            elif "WHOLE FOODS" in row[1]:
                to_account = "Expenses:Groceries"
            elif "YOUR CASH REWARD" in row[1]:
                to_account = "Income:Credit Card Rewards"
            else:
                to_account = "Expenses:Other"
                review_trans = review_trans + row[0] + ", " + row[1] + ", " + row[2] + "\n"
            amount = Decimal(row[2])
            from_account = "Liabilities:Credit Cards:Amex BlueCash Everyday"
            postdate = datetime.strptime(row[0], '%m/%d/%Y')
            with mybook as book:
                USD = mybook.currencies(mnemonic="USD")
                # create transaction with core objects in one step
                trans = Transaction(post_date=postdate.date(),
                                    currency=USD,
                                    description=row[1],
                                    splits=[
                                         Split(value=amount, memo="scripted", account=mybook.accounts(fullname=to_account)),
                                         Split(value=-amount, memo="scripted", account=mybook.accounts(fullname=from_account)),
                                     ])
                book.save()
                book.flush()
                book.close()

amex_gnu = getGnuCashBalance(mybook, 'Amex')
# # REDEEM REWARDS
# click on Rewards
driver.find_element_by_xpath("/html/body/div[1]/div[1]/div/div[2]/div/div/div[4]/div/div[2]/div/div/header/div/div/div/div/div/div/div[1]/div/div[1]/div/ul/li[5]/a/span").click()
time.sleep(10)
rewards = driver.find_element_by_xpath("/html/body/div[1]/div[1]/div/div[2]/div/div[3]/div/div[3]/div/div/div[2]/div[2]/div[1]/div/div/div/div/div/div[1]/div").text.replace("$","")

ahk = AHK()

# click on Redeem for Statement Credit
driver.find_element_by_xpath("/html/body/div[1]/div[1]/div/div[2]/div/div[3]/div/div[3]/div/div/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/a").click()
try:
    # enter full rewards amount
    driver.find_element_by_xpath("/html/body/div[1]/div/div[4]/div/div/div[2]/section[1]/div/div[1]/p[2]/input[1]").send_keys(rewards)
    # click on Redeem Now
    driver.find_element_by_xpath("/html/body/div[1]/div/div[4]/div/div/div[2]/section[1]/div/div[1]/span/a").click()
    # enter email address
    driver.find_element_by_xpath("/html/body/div[1]/div/div[4]/div/div/div[2]/div/div/div/div/div/form/div[2]/input[2]").send_keys(os.environ.get('Email'))
    time.sleep(2)
    # accept Terms & Conditions
    driver.find_element_by_xpath("/html/body/div[1]/div/div[4]/div/div/div[2]/div/div/div/div/div/form/div[2]/input[2]").click()
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('space')
    # ahk.key_press('Tab')
    # ahk.key_press('Tab')
    # ahk.key_press('Space')
    #driver.find_element_by_xpath("/html/body/div[1]/div/div[4]/div/div/div[2]/div/div/div/div/div/form/div[2]/div[2]/input[1]").click()
    # click Redeem Now (again)
    driver.find_element_by_xpath("/html/body/div[1]/div/div[4]/div/div/div[2]/div/div/div/div/div/form/div[2]/input[3]").click()
except NoSuchElementException:
    exception = "caught"
# open Checking Balance Sheet
json_creds = directory + r"\Projects\Coding\Python\BankingAutomation\Resources\creds.json"
sheet = gspread.service_account(filename=json_creds).open("Checking Balance")
# convert balance from currency (string) to negative amount
balance_str = (amex.strip('$'))
balance_num = float(balance_str)
balance = balance_num * -1
# get current m1_date
today = date.today()
year = today.year
month = today.month
# switch worksheets if running in December (to next year's worksheet)
if month == 12:
    year = year + 1
worksheet = sheet.worksheet(str(year))
# update appropriate month's information
if month == 1:
    worksheet.update('K7', balance)
    worksheet.update('N7', balance)
elif month == 2:
    worksheet.update('S7', balance)
    worksheet.update('V7', balance)
elif month == 3:
    worksheet.update('C42', balance)
    worksheet.update('F42', balance)
elif month == 4:
    worksheet.update('K42', balance)
    worksheet.update('N42', balance)
elif month == 5:
    worksheet.update('S42', balance)
    worksheet.update('V42', balance)
elif month == 6:
    worksheet.update('C77', balance)
    worksheet.update('F77', balance)
elif month == 7:
    worksheet.update('K77', balance)
    worksheet.update('N77', balance)
elif month == 8:
    worksheet.update('S77', balance)
    worksheet.update('V77', balance)
elif month == 9:
    worksheet.update('C112', balance)
    worksheet.update('F112', balance)
elif month == 10:
    worksheet.update('K112', balance)
    worksheet.update('N112', balance)
elif month == 11:
    worksheet.update('S112', balance)
    worksheet.update('V112', balance)
else:
    worksheet.update('C7', balance)
    worksheet.update('F7', balance)

# Display Checking Balance spreadsheet
driver.execute_script("window.open('https://docs.google.com/spreadsheets/d/1684fQ-gW5A0uOf7s45p9tC4GiEE5s5_fjO5E7dgVI1s/edit#gid=914927265');")
# Open GnuCash if there are transactions to review
if review_trans:
    os.startfile(directory + r"\Finances\Personal Finances\Finance.gnucash")

# Display Balance
showMessage("Balances + Review", f'Amex Balance: {amex} \n' f'GnuCash Amex Balance: {amex_gnu} \n \n' f'Review transactions:\n{review_trans}')

driver.quit()