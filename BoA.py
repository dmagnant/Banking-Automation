from selenium.common.exceptions import NoSuchElementException, ElementNotInteractableException
from selenium.webdriver.common.keys import Keys
from datetime import datetime
import time
from decimal import Decimal
import csv
import os
from piecash import Transaction, Split
from Functions import setDirectory, chromeDriverAsUser, getUsername, getPassword, openGnuCashBook, showMessage, getGnuCashBalance, updateSpreadsheet, setToAccount, importGnuTransaction

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
# get current date
today = datetime.today()
year = today.year
month = today.month
stmtmonth = today.strftime("%B")
stmtyear = str(year)
transactions_csv = os.path.join(r"C:\Users\dmagn\Downloads", stmtmonth + stmtyear + "_8549.csv")
time.sleep(2)
review_trans = ""
# Set Gnucash Book
mybook = openGnuCashBook(directory, 'Finance', False, False)

review_trans = importGnuTransaction('BoA', transactions_csv, mybook, driver, directory)

# # open CSV file at the given path
# with open(transactions_csv) as csv_file:
#     csv_reader = csv.reader(csv_file, delimiter=',')
#     line_count = 0
#     for row in csv_reader:
#         # skip header line
#         if line_count == 0:
#             line_count += 1
#         else:
#             # Skip payment (already captured in Checking Balance script
#             if "BA ELECTRONIC PAYMENT" in row[2]:
#                 continue
#             else:
#                 to_account = setToAccount('BoA', row)
#                 if to_account == "Expenses:Other":
#                     review_trans = review_trans + row[0] + ", " + row[1] + ", " + "\n"
#             amount = Decimal(row[4])
#             from_account = "Liabilities:Credit Cards:BankAmericard Cash Rewards"
#             postdate = datetime.strptime(row[0], '%m/%d/%Y')
#             with mybook as book:
#                 USD = mybook.currencies(mnemonic="USD")
#                 # create transaction with core objects in one step
#                 trans = Transaction(post_date=postdate.date(),
#                                     currency=USD,
#                                     description=row[2],
#                                     splits=[
#                                          Split(value=-amount, memo="scripted", account=mybook.accounts(fullname=to_account)),
#                                          Split(value=amount, memo="scripted", account=mybook.accounts(fullname=from_account)),
#                                      ])
#                 book.save()
#                 book.flush()
# book.close()
BoA_gnu = getGnuCashBalance(mybook, 'BoA')
# # REDEEM REWARDS
# click on View/Redeem menu
driver.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/div[2]/div[2]/div/div/div[1]/div[4]/div[3]/a").click()
time.sleep(5)
#scroll down to view button
driver.execute_script("window.scrollTo(0, 300)")
time.sleep(2)
# wait for Redeem Cash Rewards button to load, click it
driver.find_element_by_id("rewardsRedeembtn").click()
# switch to last window
driver.switch_to.window(driver.window_handles[len(driver.window_handles)-1])
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

# switch worksheets if running in December (to next year's worksheet)
if month == 12:
    year = year + 1
BoA_neg = float(BoA.strip('$')) * -1
updateSpreadsheet(directory, 'Checking Balance', year, 'BoA', month, BoA_neg)
updateSpreadsheet(directory, 'Checking Balance', year, 'BoA', month, BoA_neg, True)

# Display Checking Balance spreadsheet
driver.execute_script("window.open('https://docs.google.com/spreadsheets/d/1684fQ-gW5A0uOf7s45p9tC4GiEE5s5_fjO5E7dgVI1s/edit#gid=914927265');")
# Open GnuCash if there are transactions to review
if review_trans:
    os.startfile(directory + r"\Finances\Personal Finances\Finance.gnucash")
# Display Balance
showMessage("Balances + Review", f'BoA Balance: {BoA} \n' f'GnuCash BoA Balance: {BoA_gnu} \n \n' f'Review transactions:\n{review_trans}')
driver.quit()