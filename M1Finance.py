
from selenium.common.exceptions import NoSuchElementException
from decimal import Decimal
from datetime import datetime
import time
import csv
from Functions import setDirectory, chromeDriverAsUser, getUsername, getPassword, openGnuCashBook, showMessage, getDateRange, modifyTransactionDescription, importGnuTransaction

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

    date_range = getDateRange(today, 3)

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
                description = modifyTransactionDescription(description, amount)
                amount = amount.replace("+", "")
                row = m1_date, description, amount
                # Write to csv file
                with open(m1_activity, 'a', newline='') as file:
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
    with open(gnu_m1_activity, 'r') as t1, open(m1_activity, 'r') as t2:
        fileone = t1.readlines()
        filetwo = t2.readlines()
        line_count = 0
    for line in filetwo:
        line_count += 1
        if line not in fileone:
            review_trans = importGnuTransaction('M1', m1_activity, mybook, driver, 0)

    return [m1, review_trans]