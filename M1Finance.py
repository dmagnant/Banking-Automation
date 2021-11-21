
from selenium.common.exceptions import NoSuchElementException
from decimal import Decimal
from datetime import datetime, timedelta
import time
from piecash import Transaction, Split
import csv
from Functions import setDirectory, chromeDriverAsUser, getUsername, getPassword, openGnuCashBook, showMessage

def runM1(directory, driver):
    driver.get("https://dashboard.m1finance.com/login")
    driver.maximize_window()
    # login
    # enter email
    driver.find_element_by_name("username").send_keys(getUsername(directory, 'M1 Finance'))
    # enter password
    driver.find_element_by_name("password").send_keys(getPassword(directory, 'M1 Finance'))
    # click Log In
    driver.find_element_by_xpath("//*[@id='root']/div/div/div/div[2]/div/div/form/div[4]/div/button").click()
    # handle captcha
    showMessage('CAPTCHA',"Verify captcha, then click OK")
    try: 
        # click Spend
        driver.find_element_by_xpath("//*[@id='root']/div/div/div/div[1]/div[2]/div/div[1]/nav/a[2]/div/div/span").click()
    except NoSuchElementException:
        # handle captcha
        showMessage('CAPTCHA',"Verify captcha, then click OK")
        # click Spend
        driver.find_element_by_xpath("//*[@id='root']/div/div/div/div[1]/div[2]/div/div[1]/nav/a[2]/div/div/span").click()

    m1 = driver.find_element_by_xpath("//*[@id='root']/div/div/div/div[2]/div/div[1]/div[2]/div/div[1]/div/h1").text.replace("$", "").replace(",", "")
    # get current date
    today = datetime.today()
    year = today.year
    month = today.month

    # Gather last 3 days worth of transactions
    inside_date_range = True
    current_date = datetime.today().date()
    date_range = str(current_date)
    date_range_length_in_days = 3
    day = 1
    while day <= date_range_length_in_days:
        day_before = (current_date - timedelta(days=day)).isoformat()
        date_range = date_range + day_before
        day += 1

    transaction = 1
    column = 1
    element = "//*[@id='root']/div/div/div/div[2]/div/div[2]/div/div[3]/a[" + str(transaction) + "]/div[" + str(column) + "]"
    inside_date_range = True
    m1_activity = directory + r"\Projects\Coding\Python\BankingAutomation\Resources\m1.csv"
    gnu_m1_activity = directory + r"\Projects\Coding\Python\BankingAutomation\Resources\gnu_m1.csv"
    with open(m1_activity, 'w', newline='') as file:
        file.truncate()
    with open(gnu_m1_activity, 'w', newline='') as file:
        file.truncate()
    # Set Gnucash Book
    mybook = openGnuCashBook(directory, 'Finance', False, False)
    clicked_next = False
    while inside_date_range:
        try:
            # capture m1_date
            raw_date = driver.find_element_by_xpath(element).text
            if "," in raw_date:
                try:
                    # capture Date if Mon Day, Year format (standard for prior year)
                    m1_date = datetime.strptime(driver.find_element_by_xpath(element).text, '%b %d, %Y').date()
                except ValueError:
                    # capture Date if Day of week, Mon Day format (standard for current year)
                    mod_date = datetime.strptime(raw_date, '%a, %b %d').date()
                    # adding year since its missing in M1 Finance
                    m1_date = mod_date.replace(year=year)
            else:
                # capture date if Mon Day format (only one transaction so far. REMOVE??)
                mod_date = datetime.strptime(raw_date, '%b %d').date()
                # adding year since its missing in M1 Finance
                m1_date = mod_date.replace(year=year)
            if str(m1_date) not in date_range:
                inside_date_range = False
            else:
                column += 1
                element = "//*[@id='root']/div/div/div/div[2]/div/div[2]/div/div[3]/a[" + str(transaction) + "]/div[" + str(column) + "]"
                description = driver.find_element_by_xpath(element).text
                column += 2
                element = "//*[@id='root']/div/div/div/div[2]/div/div[2]/div/div[3]/a[" + str(transaction) + "]/div[" + str(column) + "]"
                amount = driver.find_element_by_xpath(element).text.replace("$", "").replace(",", "")
                if "Transfer to M1 Invest" in description:
                    description = "IRA Transfer"
                elif "Interest paid to M1 Spend Plus" in description:
                    description = "Interest paid"
                elif "NORTHWESTERN MUT" in description:
                    description = "NM Paycheck"
                elif "PAYPAL" in description and "10.00" in amount:
                    description = "Swagbucks"  
                elif "PAYPAL" in description:
                    description = "Paypal"
                elif "LENDING CLUB" in description:
                    description = "Lending Club"
                elif "NIELSEN" in description and "3.00" in amount:
                    description = "Pinecone Research"
                elif "VENMO" in description:
                    description = "Venmo"
                elif "TIAA" in description:
                    description = "TIAA Transfer"
                elif "Transfer from linked bank" in description:
                    description = "TIAA Transfer"
                elif "AMEX EPAYMENT" in description:
                    description = "Amex CC"
                elif "CHASE CREDIT CRD RWRD" in description:
                    description = "Chase CC Rewards"
                elif "CHASE CREDIT CRD AUTOPAY" in description:
                    description = "Chase CC"
                elif "DISCOVER E-PAYMENT" in description:
                    description = "Discover CC"
                elif "DISCOVER CASH AWARD" in description:
                    description = "Discover CC Rewards"
                elif "BARCLAYCARD US CREDITCARD" in description:
                    description = "Barclays CC"
                elif "BARCLAYCARD US ACH REWARD" in description:
                    description = "Barclays CC rewards"
                elif "BK OF AMER VISA ONLINE PMT" in description:
                    description = "BoA CC"
                elif "ALLY BANK $TRANSFER" in description:
                    description = "Ally Transfer"
                amount = amount.replace("+", "")
                row = m1_date, description, amount
                # Write to csv file
                with open(m1, 'a', newline='') as file:
                    csv_writer = csv.writer(file)
                    csv_writer.writerow(row)
                    transaction += 1
                    column = 1
                    element = "//*[@id='root']/div/div/div/div[2]/div/div[2]/div/div[3]/a[" + str(
                        transaction) + "]/div[" + str(column) + "]"
        except NoSuchElementException:
            if not clicked_next:
                # click Next
                driver.find_element_by_xpath(
                    "//*[@id='root']/div/div/div/div[2]/div/div[2]/div/div[2]/div/button[2]/span[2]").click()
                transaction = 1
                column = 1
                element = "//*[@id='root']/div/div/div/div[2]/div/div[2]/div/div[3]/a[" + str(transaction) + "]/div[" + str(
                    column) + "]"
                time.sleep(2)
                clicked_next = True
            else:
                inside_date_range = False

    account = "Assets:Liquid Assets:M1 Spend"
    # retrieve transactions from GnuCash
    transactions = [tr for tr in mybook.transactions
                    if str(tr.post_date.strftime('%Y-%m-%d')) in date_range
                    for spl in tr.splits
                    if spl.account.fullname == account]
    for tr in transactions:
        m1_date = str(tr.post_date.strftime('%Y-%m-%d'))
        description = str(tr.description)
        for spl in tr.splits:
            amount = format(spl.value, ".2f")
            if spl.account.fullname == account:
                # open CSV file at the given path
                rows = m1_date, description, amount
                with open(gnu_m1_activity, 'a', newline='') as file:
                    csv_writer = csv.writer(file)
                    csv_writer.writerow(rows)
    review_trans = ""
    with open(gnu_m1_activity, 'r') as t1, open(m1, 'r') as t2:
        fileone = t1.readlines()
        filetwo = t2.readlines()
        line_count = 0
    for line in filetwo:
        line_count += 1
        if line not in fileone:
            csv_reader = csv.reader(gnu_m1_activity, delimiter=',')
            row_count = 0
            with open(m1_activity) as file:
                csv_reader = csv.reader(file, delimiter=',')
                for row in csv_reader:
                    row_count += 1
                    if line_count == row_count:
                        if "Swagbucks" in row[1]:
                            to_account = "Income:Market Research"
                        elif "Interest paid" in row[1]:
                            to_account = "Income:Investments:Interest"
                        elif "NM Paycheck" in row[1]:
                            to_account = "Income:Salary"
                        elif "TIAA Transfer" in row[1]:
                            to_account = "Assets:Liquid Assets:TIAA"
                        elif "coinbase" in row[1].lower():
                            to_account = "Assets:Non-Liquid Assets:CryptoCurrency"
                        elif "Pinecone Research" in row[1]:
                            to_account = "Income:Market Research"
                        elif "IRA Transfer" in row[1]:
                            to_account = "Assets:Non-Liquid Assets:Roth IRA"
                        elif "Lending Club" in row[1]:
                            to_account = "Assets:Non-Liquid Assets:MicroLoans"
                        elif "Chase CC Rewards" in row[1]:
                            to_account = "Income:Credit Card Rewards"
                        elif "Chase CC" in row[1]:
                            to_account = "Liabilities:Credit Cards:Chase Freedom"
                        elif "Discover CC Rewards" in row[1]:
                            to_account = "Income:Credit Card Rewards"
                        elif "Discover CC" in row[1]:
                            to_account = "Liabilities:Credit Cards:Discover It"
                        elif "Amex CC" in row[1]:
                            to_account = "Liabilities:Credit Cards:Amex BlueCash Everyday"
                        elif "BoA CC" in row[1]:
                            to_account = "Liabilities:Credit Cards:BankAmericard Cash Rewards"
                        elif "Barclays CC Rewards" in row[1]:
                            to_account = "Income:Credit Card Rewards"
                        elif "Barclays CC" in row[1]:
                            to_account = "Liabilities:Credit Cards:BarclayCard CashForward"
                        elif "Ally Transfer" in row[1]:
                            to_account = "Expenses:Joint Expenses"
                        else:
                            to_account = "Expenses:Other"
                            review_trans = review_trans + row[0] + ", " + row[1] + ", " + row[2] + "\n"
                        amount = Decimal(row[2])
                        from_account = "Assets:Liquid Assets:M1 Spend"
                        postdate = datetime.strptime(row[0], '%Y-%m-%d')
                        with mybook as book:
                            USD = mybook.currencies(mnemonic="USD")
                            if "NM Paycheck" in row[1]:
                                review_trans = review_trans + row[0] + ", " + row[1] + ", " + row[2] + "\n"
                                entry = Transaction(post_date=postdate.date(),
                                                    currency=USD,
                                                    description=row[1],
                                                    splits=[
                                                        Split(value=round(Decimal(1871.40), 2), memo="scripted",
                                                            account=mybook.accounts(fullname=from_account)),
                                                        Split(value=round(Decimal(173.36), 2), memo="scripted",
                                                            account=mybook.accounts(
                                                                fullname="Assets:Non-Liquid Assets:401k")),
                                                        Split(value=round(Decimal(5.49), 2), memo="scripted",
                                                            account=mybook.accounts(fullname="Expenses:Medical:Dental")),
                                                        Split(value=round(Decimal(36.22), 2), memo="scripted",
                                                            account=mybook.accounts(fullname="Expenses:Medical:Health")),
                                                        Split(value=round(Decimal(2.67), 2), memo="scripted",
                                                            account=mybook.accounts(fullname="Expenses:Medical:Vision")),
                                                        Split(value=round(Decimal(168.54), 2), memo="scripted",
                                                            account=mybook.accounts(
                                                                fullname="Expenses:Income Taxes:Social Security")),
                                                        Split(value=round(Decimal(39.42), 2), memo="scripted",
                                                            account=mybook.accounts(
                                                                fullname="Expenses:Income Taxes:Medicare")),
                                                        Split(value=round(Decimal(305.08), 2), memo="scripted",
                                                            account=mybook.accounts(
                                                                fullname="Expenses:Income Taxes:Federal Tax")),
                                                        Split(value=round(Decimal(157.03), 2), memo="scripted",
                                                            account=mybook.accounts(
                                                                fullname="Expenses:Income Taxes:State Tax")),
                                                        Split(value=round(Decimal(130.00), 2), memo="scripted",
                                                            account=mybook.accounts(
                                                                fullname="Assets:Non-Liquid Assets:HSA")),
                                                        Split(value=-round(Decimal(2889.21), 2), memo="scripted",
                                                            account=mybook.accounts(fullname=to_account)),
                                                    ])
                                book.save()
                                book.flush()
                            else:
                                entry = Transaction(post_date=postdate.date(),
                                                    currency=USD,
                                                    description=row[1],
                                                    splits=[
                                                        Split(value=-amount, memo="scripted",
                                                            account=mybook.accounts(fullname=to_account)),
                                                        Split(value=amount, memo="scripted",
                                                            account=mybook.accounts(fullname=from_account)),
                                                    ])
                                book.save()
                                book.flush()
                        book.close()
    return [m1, review_trans]