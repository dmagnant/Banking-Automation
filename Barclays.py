from datetime import date, datetime
import time
from typing import KeysView
import gspread
from selenium.common.exceptions import NoSuchElementException
from decimal import Decimal
import csv
from piecash import Transaction, Split
import os
from Functions import setDirectory, chromeDriverAsUser, getUsername, getPassword, openGnuCashBook, showMessage, getGnuCashBalance

directory = setDirectory()
driver = chromeDriverAsUser(directory)
driver.implicitly_wait(3)
driver.get("https://www.barclaycardus.com/servicing/home?secureLogin=")
driver.maximize_window()

# Login
driver.find_element_by_name("uxLoginForm.username").send_keys(getUsername(directory, 'Barclay Card'))
driver.find_element_by_name("uxLoginForm.password").send_keys(getPassword(directory, 'Barclay Card'))
driver.find_element_by_name("login").click()

# handle security questions
try:
    question1 = driver.find_element_by_id("question1TxtbxLabel").text
    question2 = driver.find_element_by_id("question2TxtbxLabel").text
    q1 = "In what year was your mother born?"
    a1 = os.environ.get('MotherBirthYear')
    q2 = "What is the name of the first company you worked for?"
    a2 = os.environ.get('FirstEmployer')
    if question1 == q1:
        driver.find_element_by_id("rsaAns1").send_keys(a1)
        driver.find_element_by_id("rsaAns2").send_keys(a2)
    else:
        driver.find_element_by_id("rsaAns1").send_keys(a2)
        driver.find_element_by_id("rsaAns2").send_keys(a1)
    driver.find_element_by_id("rsaChallengeFormSubmitButton").click()
except NoSuchElementException:
    exception = "Caught"
# handle Confirm Your Identity
try:
    driver.find_element_by_xpath("/html/body/section[2]/div[4]/div/div/div[2]/form/div/div[1]/div[2]/div/div/div[1]/div/div/div[2]/div[1]/label/span").click()
    driver.find_element_by_xpath("//button[@type='submit']").click()
    # User pop-up to enter SecurPass Code
    showMessage("Get Code From Phone", "Enter in box, then click OK")
    driver.find_element_by_xpath("//button[@type='submit']").click()
except NoSuchElementException:
    exception = "Caught"
# handle Pop-up
try:
    driver.find_element_by_xpath("/html/body/div[4]/div/button/span").click()
except NoSuchElementException:
    exception = "Caught"
# # Capture Statement balance
barclays = driver.find_element_by_xpath("/html/body/section[2]/div[4]/div[2]/div[1]/section[2]/div/div/div[2]/div/div[2]/div[2]/div[2]/div[2]").text.replace('-', '')
# Capture Rewards balance
rewards_balance = driver.find_element_by_xpath("//*[@id='rewardsTile']/div[2]/div/div[2]/div[1]/div").text.replace('$', '')
# # EXPORT TRANSACTIONS
# click on Activity & Statements
driver.find_element_by_xpath("/html/body/section[2]/div[1]/nav/div/ul/li[3]/a").click()
# Click on Transactions
driver.find_element_by_xpath("/html/body/section[2]/div[1]/nav/div/ul/li[3]/ul/li/div/div[2]/ul/li[1]/a").click()
# Click on Download
driver.find_element_by_xpath("/html/body/section[2]/div[4]/div/div/div[3]/div[1]/div/div[2]/span/div/button/span[1]").click()
# determine proper date_range
today = date.today()
month = today.month
monthto = str(month)
year = today.year
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
filename = r"C:\Users\dmagn\Downloads\CreditCard_" + yearfrom + monthfrom + "11_" + yearto + monthto + "10.csv"
# enter date_range
driver.find_element_by_id("downloadFromDate_input").send_keys(fromdate)
driver.find_element_by_id("downloadToDate_input").send_keys(todate)
# click Download
driver.find_element_by_xpath("/html/body/div[3]/div[2]/div/div/div[2]/div/form/div[3]/div/button").click()
# # IMPORT TRANSACTIONS
review_trans = ""
# open CSV file at the given path
time.sleep(2)
# Set Gnucash Book
mybook = openGnuCashBook(directory, 'Finance', False, False)
with open(filename) as csv_file:
    csv_reader = csv.reader(csv_file, delimiter=',')
    line_count = 0
    for row in csv_reader:
        # skip header line
        if line_count < 5:
            line_count += 1
        else:
            # Skip payment (already captured in Checking Balance script)
            if "Payment Received" in row[1]:
                skip = "yes"
            elif "SPECTRUM" in row[1]:
                to_account = "Expenses:Housing:Internet"
            elif "google fi" in row[1].lower():
                to_account = "Expenses:Cell Phone"
            elif "CAT DOCTOR" in row[1]:
                to_account = "Expenses:Medical:Vet"
            elif "PARKING" in row[1]:
                to_account = "Expenses:Transportation:Parking"
            elif "PICK N SAVE" in row[1]:
                to_account = "Expenses:Groceries"
            elif "KETTLE RANGE" in row[1]:
                to_account = "Expenses:Groceries"
            elif "KOPPA" in row[1]:
                to_account = "Expenses:Groceries"
            elif "AMZN" in row[1]:
                to_account = "Expenses:Amazon"
            elif "Amazon" in row[1]:
                to_account = "Expenses:Amazon"
            elif "STEAMGAMES" in row[1]:
                to_account = "Expenses:Entertainment"
            elif "STEAM PURCHASE" in row[1]:
                to_account = "Expenses:Entertainment"
            elif "PROGRESSIVE" in row[1]:
                to_account = "Expenses:Transportation:Car Insurance"
            elif "UBER" in row[1]:
                to_account = "Expenses:Transportation:Ride Services"
            else:
                to_account = "Expenses:Other"
                review_trans = review_trans + row[0] + ", " + row[1] + ", " + row[3] + "\n"
            amount = Decimal(row[3])
            from_account = "Liabilities:Credit Cards:BarclayCard CashForward"
            postdate = datetime.strptime(row[0], '%m/%d/%Y')
            
            with mybook as book:
                USD = mybook.currencies(mnemonic="USD")
                # create transaction with core objects in one step
                if "Payment Received" in row[1]:
                    continue
                else:
                    trans = Transaction(post_date=postdate.date(),
                                        currency=USD,
                                        description=row[1],
                                        splits=[
                                            Split(value=-amount, memo="scripted", account=mybook.accounts(fullname=to_account)),
                                            Split(value=amount, memo="scripted", account=mybook.accounts(fullname=from_account)),
                                        ])
                book.save()
                book.flush()
barclays_gnu = getGnuCashBalance(mybook, 'Barclays')
if float(rewards_balance) > 50:
    # # REDEEM REWARDS
    barclays_window = driver.window_handles[0]
    driver.switch_to.window(barclays_window)
    # click on Rewards & Benefits
    driver.find_element_by_xpath("/html/body/section[2]/div[1]/nav/div/ul/li[4]/a").click()
    # click on Redeem my cash rewards
    driver.find_element_by_xpath("//*[@id='rewards-benefits-container']/div[1]/ul/li[3]/a").click()
    # click on Direct Deposit or Statement Credit
    driver.find_element_by_xpath("/html/body/div[1]/div[2]/div[2]/div[1]/div/div/div[2]/ul/li[1]/div/a").click()
    # click on Continue
    driver.find_element_by_id("redeem-continue").click()
    # Click "select an option" (for method of rewards)
    driver.find_element_by_xpath("//*[@id='mor_dropDown0']").click()
    driver.find_element_by_xpath("//*[@id='mor_dropDown0']").send_keys(KeysView.DOWN)
    driver.find_element_by_xpath("//*[@id='mor_dropDown0']").send_keys(KeysView.ENTER)
    time.sleep(1)
    # Click Continue
    driver.find_element_by_xpath("//*[@id='achModal-continue']").click()
    time.sleep(1)
    # click on Redeem Now
    driver.find_element_by_xpath("/html/body/section[2]/div[4]/div[2]/cashback/div/div[2]/div/ui-view/redeem/div/review/div/div/div/div/div[2]/form/div[3]/div/div[1]/button").click()

# open Checking Balance Sheet
json_creds = directory + r"\Projects\Coding\Python\BankingAutomation\Resources\creds.json"
sheet = gspread.service_account(filename=json_creds).open("Checking Balance")
# convert balance from currency (string) to negative amount
balance_str = (barclays.strip('$'))
balance_num = balance_str.replace(',', '').replace('-','')
balance_inv = float(balance_num)
balance = balance_inv * -1
# get current m1_date
year = today.year
month = today.month
# switch worksheets if running in December (to next year's worksheet)
if month == 12:
    year = year + 1
worksheet = sheet.worksheet(str(year))
# update appropriate month's information
if month == 1:
    worksheet.update('K10', balance)
    worksheet.update('N10', balance)
elif month == 2:
    worksheet.update('S10', balance)
    worksheet.update('V10', balance)
elif month == 3:
    worksheet.update('C45', balance)
    worksheet.update('F45', balance)
elif month == 4:
    worksheet.update('K45', balance)
    worksheet.update('N45', balance)
elif month == 5:
    worksheet.update('S45', balance)
    worksheet.update('V45', balance)
elif month == 6:
    worksheet.update('C80', balance)
    worksheet.update('F80', balance)
elif month == 7:
    worksheet.update('K80', balance)
    worksheet.update('N80', balance)
elif month == 8:
    worksheet.update('S80', balance)
    worksheet.update('V80', balance)
elif month == 9:
    worksheet.update('C115', balance)
    worksheet.update('F115', balance)
elif month == 10:
    worksheet.update('K115', balance)
    worksheet.update('N115', balance)
elif month == 11:
    worksheet.update('S115', balance)
    worksheet.update('V115', balance)
else:
    worksheet.update('C10', balance)
    worksheet.update('F10', balance)
# Display Checking Balance spreadsheet
driver.execute_script("window.open('https://docs.google.com/spreadsheets/d/1684fQ-gW5A0uOf7s45p9tC4GiEE5s5_fjO5E7dgVI1s/edit#gid=914927265');")
# Open GnuCash if there are transactions to review
if review_trans:
    os.startfile(directory + r"\Finances\Personal Finances\Finance.gnucash")
# Display Balance
showMessage("Balances + Review", f'Barclays balance: {barclays} \n' f'GnuCash Barclays balance: {barclays_gnu} \n \n' f'Review transactions:\n{review_trans}')
driver.quit()