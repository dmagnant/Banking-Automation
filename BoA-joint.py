from selenium.common.exceptions import NoSuchElementException
from datetime import datetime
import time
from decimal import Decimal
import csv
import os
from piecash import Transaction, Split
from Functions import setDirectory, chromeDriverAsUser, getUsername, getPassword, openGnuCashBook, showMessage, getGnuCashBalance, updateSpreadsheet

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

driver.find_element_by_partial_link_text("Bank of America Travel Rewards Visa Signature - 8955").click()
# # Capture Statement balance
BoAjoint = driver.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/div[2]/div[2]/div/div/div[3]/div[4]/div[3]/div/div[2]/div[2]/div[2]").text
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
# modify file for import
stmtmonth = today.strftime("%B")
stmtyear = str(year)
filename = os.path.join(r"C:\Users\dmagn\Downloads", stmtmonth + stmtyear + "_8955.csv")
time.sleep(2)
review_trans = ""
energy_bill_num = 0
# Set Gnucash Book
mybook = openGnuCashBook(directory, 'Home', False, False)
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
            elif "SPECTRUM" in row[2]:
                to_account = "Expenses:Utilities:Internet"
            elif "MCDONALD'S" in row[2]:
                to_account = "Expenses:Bars & Restaurants"
            elif "JIMMY JOHNS" in row[2]:
                to_account = "Expenses:Bars & Restaurants"
            elif "EATSTREET" in row[2]:
                to_account = "Expenses:Bars & Restaurants"
            elif "GRUBHUB" in row[2]:
                to_account = "Expenses:Bars & Restaurants"
            elif "LAKEFRONT BREWERY" in row[2]:
                to_account = "Expenses:Bars & Restaurants"
            elif "CAFE CORAZON" in row [2]:
                to_account = "Expenses:Bars & Restaurants"
            elif "CAFE CORAZON" in row [2]:
                to_account = "Expenses:Bars & Restaurants"               
            elif "THE HOME DEPOT" in row[2]:
                to_account = "Expenses:Home Depot"
            elif "HOMEDEPOT.COM" in row[2]:
                to_account = "Expenses:Home Depot"
            elif "PICK N SAVE" in row[2]:
                to_account = "Expenses:Groceries"
            elif "TARGET" in row[2]:
                to_account = "Expenses:Groceries"
            elif "KETTLE RANGE" in row[2]:
                to_account = "Expenses:Groceries"
            elif "GOOGLE FI" in row[2].upper():
                to_account = "Expenses:Utilities:Phone"
            elif "AMZN Mktp" in row[2]:
                to_account = "Expenses:Amazon"
            elif "AMAZON" in row[2]:
                to_account = "Expenses:Amazon"
            elif "Amazon" in row[2]:
                to_account = "Expenses:Amazon"
            elif "UBER" in row[2]:
                to_account = "Expenses:Travel:Ride Services"
            elif "WAYFAIR" in row[2]:
                to_account = "Expenses:Home Furnishings"
            elif "TRAVELLING BEER GARDEN" in row[2]:
                to_account = "Expenses:Bars & Restaurants"
            elif "MILW COUNTY PARKS" in row[2]:
                to_account = "Expenses:Bars & Restaurants"
            elif "BOTANAS" in row[2]:
                to_account = "Expenses:Bars & Restaurants"
            elif "CHEWY" in row[2]:
                to_account = "Expenses:Pet"
            else:
                to_account = "Expenses:Other"
                review_trans = review_trans + row[0] + ", " + row[2] + ", " + row[4] + "\n"
            amount = Decimal(row[4])
            from_account = "Liabilities:BoA Credit Card"
            postdate = datetime.strptime(row[0], '%m/%d/%Y')
            with mybook as book:
                USD = mybook.currencies(mnemonic="USD")
                # create transaction with core objects in one step
                if "ARCADIA" in row[2]:
                    energy_bill_num += 1
                    if energy_bill_num == 1:
                        # Get balances from Arcadia
                        driver.execute_script("window.open('https://login.arcadia.com/email');")
                        driver.implicitly_wait(5)
                        # make the new window active
                        arcadia_window = driver.window_handles[1]
                        driver.switch_to.window(arcadia_window)
                        # get around bot-prevention by logging in twice
                        num = 1
                        while num <3:
                            try:
                                # click Sign in with email
                                driver.find_element_by_xpath("/html/body/div/main/div[1]/div/div/div[1]/div/a").click()
                                time.sleep(1)
                            except NoSuchElementException:
                                exception = "sign in page loaded already"
                            try:
                                # Login
                                driver.find_element_by_xpath("/html/body/div[1]/main/div[1]/div/form/div[1]/div[1]/input").send_keys(getUsername(directory, 'Arcadia Power'))
                                time.sleep(1)
                                driver.find_element_by_xpath("/html/body/div[1]/main/div[1]/div/form/div[1]/div[2]/input").send_keys(getPassword(directory, 'Arcadia Power'))
                                time.sleep(1)
                                driver.find_element_by_xpath("/html/body/div[1]/main/div[1]/div/form/div[2]/button").click()
                                time.sleep(1)
                                # Get Billing page
                                driver.get("https://home.arcadia.com/billing")
                            except NoSuchElementException:
                                exception = "already signed in"
                            num += 1
                        showMessage("Login Check", 'Confirm Login to Arcadia, (manually if necessary) \n' 'Then click OK \n')
                    else:
                        driver.switch_to.window(arcadia_window)
                        driver.find_element_by_xpath("https://home.arcadia.com/billing")
                    statement_row = 1
                    statement_found = "no"
                    while statement_found == "no":
                        # Capture statement balance
                        arcadia_balance = driver.find_element_by_xpath("/html/body/div[1]/div[2]/div/div/div[3]/div/div[2]/ul/li[" + str(statement_row) + "]/div/div/p").text.replace('$', '')
                        formatted_amount = "{:.2f}".format(abs(amount))
                        if arcadia_balance == formatted_amount:
                            # click to view statement
                            driver.find_element_by_xpath("/html/body/div[1]/div[2]/div/div/div[3]/div/div[2]/ul/li[" + str(statement_row) + "]/div/div/p").click()
                            statement_found = "yes"
                        else:
                            statement_row += 1

                    # comb through lines of Arcadia Statement for Arcadia Membership (and Free trial rebate), Community Solar lines (3)
                    arcadia_statement_lines_left = True
                    statement_row = 1
                    solar = 0
                    arcadia_membership = 0
                    while arcadia_statement_lines_left:
                        try:
                            # read the header to get transaction description
                            statement_trans = driver.find_element_by_xpath("/html/body/div[1]/div[2]/div[2]/div[5]/ul/li[" + str(statement_row) + "]/div/h2").text
                            if statement_trans == "Arcadia Membership":
                                arcadia_membership = Decimal(driver.find_element_by_xpath("/html/body/div[1]/div[2]/div[2]/div[5]/ul/li[" + str(statement_row) + "]/div/p").text.replace('$',''))
                                arcadiaamt = Decimal(arcadia_membership)
                            elif statement_trans == "Free Trial":
                                arcadia_membership = arcadia_membership + Decimal(driver.find_element_by_xpath("/html/body/div[1]/div[2]/div[2]/div[5]/ul/li[" + str(statement_row) + "]/div/p").text.replace('$',''))
                            elif statement_trans == "Community Solar":
                                solar = solar + Decimal(driver.find_element_by_xpath("/html/body/div[1]/div[2]/div[2]/div[5]/ul/li[" + str(statement_row) + "]/div/p").text.replace('$',''))
                            elif statement_trans == "WE Energies Utility":
                                we_bill = driver.find_element_by_xpath("/html/body/div[1]/div[2]/div[2]/div[5]/ul/li[" + str(statement_row) + "]/div/p").text
                            statement_row += 1
                        except NoSuchElementException:
                            arcadia_statement_lines_left = False
                    
                    arcadiaamt = Decimal(arcadia_membership)
                    solaramt = Decimal(solar)

                    # Get balances from WE Energies
                    if energy_bill_num == 1:
                        driver.execute_script("window.open('https://www.we-energies.com/secure/auth/l/acct/summary_accounts.aspx');")
                        # make the new window active
                        we_window = driver.window_handles[2]
                        driver.switch_to.window(we_window)
                        try:
                            ## LOGIN
                            driver.find_element_by_xpath("//*[@id='signInName']").send_keys(getUsername(directory, 'WE-Energies (Home)'))
                            driver.find_element_by_xpath("//*[@id='password']").send_keys(getPassword(directory, 'WE-Energies (Home)'))
                            # click Login
                            driver.find_element_by_xpath("//*[@id='next']").click()
                            time.sleep(4)
                            # close out of app notice
                            driver.find_element_by_xpath("//*[@id='notInterested']/a").click
                        except NoSuchElementException:
                            exception = "caught"
                        # Click View bill history
                        driver.find_element_by_xpath("//*[@id='mainContentCopyInner']/ul/li[2]/a").click()
                        time.sleep(4)
                    bill_row = 2
                    bill_column = 7
                    bill_found = "no"
                    # find bill based on comparing amount from Arcadia (we_bill)
                    while bill_found == "no":
                        # capture date
                        we_bill_path = "/html/body/div[1]/div[1]/form/div[5]/div/div/div/div/div[6]/div[2]/div[2]/div/table/tbody/tr[" + str(bill_row) + "]/td[" + str(bill_column) + "]/span/span"
                        we_bill_amount = driver.find_element_by_xpath(we_bill_path).text
                        if we_bill == we_bill_amount:
                            bill_found = "yes"
                        else:
                            bill_row += 1
                    # capture gas charges
                    bill_column -= 2
                    we_amt_path = "/html/body/div[1]/div[1]/form/div[5]/div/div/div/div/div[6]/div[2]/div[2]/div/table/tbody/tr[" + str(bill_row) + "]/td[" + str(bill_column) + "]/span"
                    gasamt = Decimal(driver.find_element_by_xpath(we_amt_path).text.replace('$', ""))
                    # capture electricity charges
                    bill_column -= 2
                    we_amt_path = "/html/body/div[1]/div[1]/form/div[5]/div/div/div/div/div[6]/div[2]/div[2]/div/table/tbody/tr[" + str(bill_row) + "]/td[" + str(bill_column) + "]/span"
                    electricityamt = Decimal(driver.find_element_by_xpath(we_amt_path).text.replace('$', ""))
                    trans = Transaction(post_date=postdate.date(),
                                        currency=USD,
                                        description=row[2],
                                        splits=[
                                            Split(value=arcadiaamt, memo="Arcadia Membership Fee", account=mybook.accounts(fullname="Expenses:Utilities:Arcadia Membership")),
                                            Split(value=solaramt, memo="Solar Rebate", account=mybook.accounts(fullname="Expenses:Utilities:Arcadia Membership")),
                                            Split(value=electricityamt, account=mybook.accounts(fullname="Expenses:Utilities:Electricity")),
                                            Split(value=gasamt, account=mybook.accounts(fullname="Expenses:Utilities:Gas")),
                                            Split(value=amount, account=mybook.accounts(fullname=from_account)),
                                        ])
                else:
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
BoAjoint_gnu = getGnuCashBalance(mybook, 'BoA-joint')
# switch worksheets if running in December (to next year's worksheet)
if month == 12:
    year = year + 1
BoAjoint_neg = float(BoAjoint.strip('$')) * -1
updateSpreadsheet(directory, 'Home', year, 'BoA-joint', month, BoAjoint_neg)
updateSpreadsheet(directory, 'Home', year, 'BoA-joint', month, BoAjoint_neg, True)
# Display Home spreadsheet
driver.execute_script("window.open('https://docs.google.com/spreadsheets/d/1oP3U7y8qywvXG9U_zYXgjFfqHrCyPtUDl4zPDftFCdM/edit#gid=460564976');")
# Open GnuCash if there are transactions to review
if review_trans:
    os.startfile(directory + r"\Stuff\Home\Finances\Home.gnucash")
# Display Balance
showMessage("Balances + Review", f'BoA Balance: {BoAjoint} \n'f'GnuCash BoA Balance: {BoAjoint_gnu} \n \n' f'Review transactions:\n{review_trans}')
driver.quit()