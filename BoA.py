from selenium.common.exceptions import NoSuchElementException, ElementNotInteractableException
from selenium.webdriver.common.keys import Keys
import gspread
from datetime import date, datetime
import time
from decimal import Decimal
import csv
import os
from piecash import Transaction, Split
from Functions import setDirectory, chromeDriverAsUser, getUsername, getPassword, openGnuCashBook, showMessage, getGnuCashBalance

directory = setDirectory()
driver = chromeDriverAsUser(directory)
driver.implicitly_wait(3)
driver.get("https://www.bankofamerica.com/")
driver.maximize_window()
# login
driver.find_element_by_id("onlineId1").send_keys(getUsername(directory, 'BoA CC'))
driver.find_element_by_id("passcode1").send_keys(getPassword(directory, 'BoA CC'))
driver.find_element_by_xpath("//*[@id='signIn']").click()
# handle ID verification
try:
    driver.find_element_by_xpath("//*[@id='btnARContinue']/span[1]").click()
    showMessage("Get Verification Code", "Enter code, then click OK")
    driver.find_element_by_xpath("//*[@id='yes-recognize']").click()
    driver.find_element_by_xpath("//*[@id='continue-auth-number']/span").click()
except NoSuchElementException:
    exception = "Caught"

# handle security questions
try:
    question = driver.find_element_by_xpath("/html/body/div[1]/div/div/div[2]/div[1]/div/div/form/div[2]/label").text
    q1 = "What is the name of a college you applied to but didn't attend?"
    a1 = os.environ.get('CollegeApplied')
    q2 = "As a child, what did you want to be when you grew up?"
    a2 = os.environ.get('DreamJob')
    q3 = "What is the name of your first babysitter?"
    a3 = os.environ.get('FirstBabySitter')
    if driver.find_element_by_xpath("/html/body/div[1]/div/div/div[2]/div[1]/div/div/form/div[2]/label"):
        question = driver.find_element_by_xpath(
            "/html/body/div[1]/div/div/div[2]/div[1]/div/div/form/div[2]/label").text
        if question == q1:
            driver.find_element_by_name("challengeQuestionAnswer").send_keys(a1)
        elif question == q2:
            driver.find_element_by_name("challengeQuestionAnswer").send_keys(a2)
        else:
            driver.find_element_by_name("challengeQuestionAnswer").send_keys(a3)
        driver.find_element_by_xpath("/html/body/div[1]/div/div/div[2]/div[1]/div/div/form/fieldset/div[2]/div/div[1]/input").click()
        driver.find_element_by_xpath("/html/body/div[1]/div/div/div[2]/div[1]/div/div/form/a[1]/span").click()
except NoSuchElementException:
    exception = "Caught"

# close mobile app pop-up
try:
    driver.find_element_by_xpath("//*[@id='sasi-overlay-module-modalClose']/span[1]").click()
except NoSuchElementException:
    exception = "Caught"

driver.find_element_by_partial_link_text("Customized Cash Rewards Visa Signature - 8549").click()
# # Capture Statement balance
BoA = driver.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/div[2]/div[2]/div/div/div[3]/div[4]/div[3]/div/div[2]/div[2]/div[2]").text
# # EXPORT TRANSACTIONS
# click Previous transactions
time.sleep(3)
driver.find_element_by_partial_link_text("Previous transactions").click()
# click Download
driver.find_element_by_xpath("/html/body/div[1]/div/div[4]/div[1]/div/div[4]/div[2]/div[2]/div/div[1]/a").click()
# select Microsoft Excel format
driver.find_element_by_xpath("/html/body/div[1]/div/div[4]/div[1]/div/div[4]/div[2]/div[2]/div/div[3]/div/div[3]/div[1]/select").send_keys("m")
# click Download Transactions
driver.find_element_by_xpath("/html/body/div[1]/div/div[4]/div[1]/div/div[4]/div[2]/div[2]/div/div[3]/div/div[4]/div[2]/a/span").click()
# modify file for import
today = date.today()
year = today.year
stmtmonth = today.strftime("%B")
stmtyear = str(year)
filename = os.path.join(r"C:\Users\dmagn\Downloads", stmtmonth + stmtyear + "_8549.csv")
time.sleep(2)
review_trans = ""
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
            if "BA ELECTRONIC PAYMENT" in row[2]:
                continue
            elif "CASH REWARDS" in row[2]:
                to_account = "Income:Credit Card Rewards"
            elif "MCDONALD" in row[2]:
                to_account = "Expenses:Bars & Restaurants"
            elif "OAKLAND GYRO" in row[2]:
                to_account = "Expenses:Bars & Restaurants"
            elif "GRUBHUB" in row[2]:
                to_account = "Expenses:Bars & Restaurants"
            elif "JIMMY JOHN" in row[2]:
                to_account = "Expenses:Bars & Restaurants"
            elif "LA MASA" in row[2]:
                to_account = "Expenses:Bars & Restaurants"
            elif "GOOD LAND" in row[2]:
                to_account = "Expenses:Bars & Restaurants"
            elif "CROSSROADS COLLECTIVE" in row[2]:
                to_account = "Expenses:Bars & Restaurants"
            elif "BP#" in row[2]:
                to_account = "Expenses:Transportation:Gas (Vehicle)"
            elif "AMAZON" in row[2]:
                to_account = "Expenses:Amazon"
            elif "AMZN" in row[2]:
                to_account = "Expenses:Amazon"
            else:
                to_account = "Expenses:Other"
                review_trans = review_trans + row[0] + ", " + row[3] + ", " + row[4] + "\n"
            amount = Decimal(row[4])
            from_account = "Liabilities:Credit Cards:BankAmericard Cash Rewards"
            postdate = datetime.strptime(row[0], '%m/%d/%Y')
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
BoA_gnu = getGnuCashBalance(mybook, 'BoA')
book.close()
# # REDEEM REWARDS
# click on View/Redeem menu
driver.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/div[2]/div[2]/div/div/div[1]/div[4]/div[3]/a").click()
time.sleep(5)
#scroll down to view button
driver.execute_script("window.scrollTo(0, 300)")
time.sleep(2)
# wait for Redeem Cash Rewards button to load, click it
driver.find_element_by_id("rewardsRedeembtn").click()
# make the new window active
window_after = driver.window_handles[1]
driver.switch_to.window(window_after)
# close out of pop-up (if present)
try:
    driver.find_element_by_xpath("/html/body/div[1]/div/div/div[2]/div/div/button").click()
except NoSuchElementException:
    exception = "caught"
# if no pop-up, proceed to click on Redemption option
driver.find_element_by_id("redemption_option").click()
# redeem if there is a balance, else skip
try:
    # Choose Visa - statement credit
    driver.find_element_by_id("redemption_option").send_keys("v")
    driver.find_element_by_id("redemption_option").send_keys(Keys.ENTER)
    # click on Redeem all
    driver.find_element_by_xpath("/html/body/main/div[2]/div/div/div/div/div[1]/div[2]/div[1]/div/div/form/button").click()
    # click Complete Redemption
    driver.find_element_by_xpath("/html/body/div[1]/div/div/div[3]/button[1]").click()
except ElementNotInteractableException:
    exception = "caught"
# open Checking Balance Sheet
json_creds = directory + r"\Projects\Coding\Python\BankingAutomation\Resources\creds.json"
sheet = gspread.service_account(filename=json_creds).open("Checking Balance")
# convert balance from currency (string) to negative amount
balance_str = (BoA.replace("$", ""))
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
    worksheet.update('K5', balance)
    worksheet.update('N5', balance)
elif month == 2:
    worksheet.update('S5', balance)
    worksheet.update('V5', balance)
elif month == 3:
    worksheet.update('C40', balance)
    worksheet.update('F40', balance)
elif month == 4:
    worksheet.update('K40', balance)
    worksheet.update('N40', balance)
elif month == 5:
    worksheet.update('S40', balance)
    worksheet.update('V40', balance)
elif month == 6:
    worksheet.update('C75', balance)
    worksheet.update('F75', balance)
elif month == 7:
    worksheet.update('K75', balance)
    worksheet.update('N75', balance)
elif month == 8:
    worksheet.update('S75', balance)
    worksheet.update('V75', balance)
elif month == 9:
    worksheet.update('C110', balance)
    worksheet.update('F110', balance)
elif month == 10:
    worksheet.update('K110', balance)
    worksheet.update('N110', balance)
elif month == 11:
    worksheet.update('S110', balance)
    worksheet.update('V110', balance)
else:
    worksheet.update('C5', balance)
    worksheet.update('F5', balance)
# Display Checking Balance spreadsheet
driver.execute_script("window.open('https://docs.google.com/spreadsheets/d/1684fQ-gW5A0uOf7s45p9tC4GiEE5s5_fjO5E7dgVI1s/edit#gid=914927265');")
# Open GnuCash if there are transactions to review
if review_trans:
    os.startfile(directory + r"\Finances\Personal Finances\Finance.gnucash")
# Display Balance
showMessage("Balances + Review", f'BoA Balance: {BoA} \n' f'GnuCash BoA Balance: {BoA_gnu} \n \n' f'Review transactions:\n{review_trans}')
driver.quit()