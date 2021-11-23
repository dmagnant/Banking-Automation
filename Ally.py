from selenium.common.exceptions import NoSuchElementException
from decimal import Decimal
from datetime import datetime, timedelta
import time
from piecash import Transaction, Split
import csv
from Functions import getPassword, startExpressVPN, closeExpressVPN, openGnuCashBook, getDateRange, modifyTransactionDescription, setToAccount, importGnuTransaction

def runAlly(directory, driver):
    closeExpressVPN()
    driver.implicitly_wait(5)
    driver.get("https://secure.ally.com/")
    driver.maximize_window()
    time.sleep(1)
    # login
    # enter password
    driver.find_element_by_xpath("/html/body/div/div[1]/main/div/div/div/div/div[1]/form/div[2]/div/input").send_keys(getPassword(directory, 'Ally Bank'))
    # click Log In
    driver.find_element_by_xpath("/html/body/div/div[1]/main/div/div/div/div/div[1]/form/div[3]/button/span").click()
    time.sleep(4)
    # capture balance
    ally = driver.find_element_by_xpath("/html/body/div[1]/div[1]/main/div/div/div/div[2]/div/div[2]/div/table/tbody/tr/td[2]/div").text.replace("$", "").replace(",", "")
    # click Joint Checking link
    driver.find_element_by_partial_link_text("Joint Checking").click()

    # get current date
    today = datetime.today()
    year = today.year

    date_range = getDateRange(today, 5)

    table = 2
    transaction = 1
    column = 1
    element = "//*[@id='form-elements-section']/section/section/table[" + str(table) + "]/tbody/tr[" + str(transaction) + "]/td[" + str(column) + "]"

    ally_activity = directory + r"\Projects\Coding\Python\BankingAutomation\Resources\ally.csv"
    gnu_ally_activity = directory + r"\Projects\Coding\Python\BankingAutomation\Resources\gnu_ally.csv"
    with open(ally_activity, 'w', newline='') as file:
        file.truncate()
    with open(gnu_ally_activity, 'w', newline='') as file:
        file.truncate()
    inside_date_range = True
    while inside_date_range:
        try:
            mod_date = datetime.strptime(driver.find_element_by_xpath(element).text, '%b %d, %Y').date()
            if str(mod_date) not in date_range:
                inside_date_range = False
            else:
                column += 1
                element = "//*[@id='form-elements-section']/section/section/table[" + str(table) + "]/tbody/tr[" + str(transaction) + "]/td[" + str(column) + "]/button"
                description = driver.find_element_by_xpath(element).text
                column += 1
                element = "//*[@id='form-elements-section']/section/section/table[" + str(table) + "]/tbody/tr[" + str(transaction) + "]/td[" + str(column) + "]"
                amount = driver.find_element_by_xpath(element).text.replace("$", "").replace(",", "")
                description = modifyTransactionDescription(description)
                row = str(mod_date), description, amount
                # Write to csv file
                with open(ally_activity, 'a', newline='') as file:
                    csv_writer = csv.writer(file)
                    csv_writer.writerow(row)
                    transaction += 2
                    column = 1
                    element = "//*[@id='form-elements-section']/section/section/table[" + str(table) + "]/tbody/tr[" + str(transaction) + "]/td[" + str(column) + "]"
        except NoSuchElementException:
            table += 1
            column = 1
            transaction = 1
            element = "//*[@id='form-elements-section']/section/section/table[" + str(table) + "]/tbody/tr[" + str(transaction) + "]/td[" + str(column) + "]"

    # Set Gnucash Book
    mybook = openGnuCashBook(directory, 'Home', False, False)

    account = "Assets:Ally Checking Account"
    # retrieve transactions from GnuCash
    transactions = [tr for tr in mybook.transactions
                    if str(tr.post_date.strftime('%Y-%m-%d')) in date_range
                    for spl in tr.splits
                    if spl.account.fullname == account]
    for tr in transactions:
        date = str(tr.post_date.strftime('%Y-%m-%d'))
        description = str(tr.description)
        for spl in tr.splits:
            amount = format(spl.value, ".2f")
            if spl.account.fullname == account:
                # open CSV file at the given path
                rows = date, description, amount
                with open(gnu_ally_activity, 'a', newline='') as file:
                    csv_writer = csv.writer(file)
                    csv_writer.writerow(rows)
    review_trans = ""
    with open(gnu_ally_activity, 'r') as t1, open(ally_activity, 'r') as t2:
        fileone = t1.readlines()
        filetwo = t2.readlines()
        line_count = 0
    for line in filetwo:
        line_count += 1
        if line not in fileone:
            csv_reader = csv.reader(gnu_ally_activity, delimiter=',')
            row_count = 0
            with open(ally_activity) as file:
                csv_reader = csv.reader(file, delimiter=',')
                for row in csv_reader:
                    row_count += 1
                    if line_count == row_count:
                        to_account = setToAccount('Ally', row)
                        if to_account == "Expenses:Other":
                            review_trans = review_trans + row[0] + ", " + row[1] + ", " + "\n"
                        amount = Decimal(row[2])
                        from_account = "Assets:Ally Checking Account"
                        postdate = datetime.strptime(row[0], '%Y-%m-%d')
                        with mybook as book:
                            USD = mybook.currencies(mnemonic="USD")
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
    return [ally, review_trans]