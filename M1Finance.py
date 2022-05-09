from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from datetime import datetime
import time
import csv
from Functions import getUsername, getPassword, openGnuCashBook, showMessage, getDateRange, modifyTransactionDescription, compileGnuTransactions, setDirectory, chromeDriverAsUser

def login(directory, driver):
    driver.get("https://dashboard.m1finance.com/login")
    driver.implicitly_wait(3)
    # login
    # enter email
    driver.find_element(By.NAME, "username").send_keys(getUsername(directory, 'M1 Finance'))
    # enter password
    driver.find_element(By.NAME, "password").send_keys(getPassword(directory, 'M1 Finance'))
    # click Log In
    driver.find_element(By.XPATH, "//*[@id='root']/div/div/div/div[2]/div/div/form/div[4]/div/button").click()
    # handle captcha
    # showMessage('CAPTCHA',"Verify captcha, then click OK")
    try: 
        # click Spend
        driver.find_element(By.XPATH, "//*[@id='root']/div/div/div/div[1]/div[2]/div/div[1]/nav/a[3]/div/div/span").click()
    except NoSuchElementException:
        # handle captcha
        showMessage('CAPTCHA',"Verify captcha, then click OK")
        # click Spend
        driver.find_element(By.XPATH, "//*[@id='root']/div/div/div/div[1]/div[2]/div/div[1]/nav/a[3]/div/div/span").click()


def captureBalance(driver):
    return driver.find_element(By.XPATH, "//*[@id='root']/div/div/div/div[2]/div/div[1]/div[2]/div/div[1]/div/h1").text.strip('$').replace(',', '')


def captureTransactions(driver, dateRange, m1Activity, today):
    year = today.year
    transaction = 1
    column = 1
    element = "//*[@id='root']/div/div/div/div[2]/div/div[2]/div/div[3]/a[" + str(transaction) + "]/div[" + str(column) + "]"
    insideDateRange = True
    
    # Set Gnucash Book
    clickedNext = False
    while insideDateRange:
        try:
            # capture m1_date
            rawDate = driver.find_element(By.XPATH, element).text
            if "," in rawDate:
                try:
                    # capture Date if Mon Day, Year format (standard for prior year)
                    m1Date = datetime.strptime(driver.find_element(By.XPATH, element).text, '%b %d, %Y').date()
                except ValueError:
                    # capture Date if Day of week, Mon Day format (standard for current year)
                    modDate = datetime.strptime(rawDate, '%a, %b %d').date()
                    # adding year since its missing in M1 Finance
                    m1Date = modDate.replace(year=year)
            else:
                # capture date if Mon Day format (only one transaction so far. REMOVE??)
                modDate = datetime.strptime(rawDate, '%b %d').date()
                # adding year since its missing in M1 Finance
                m1Date = modDate.replace(year=year)
            if str(m1Date) not in dateRange:
                insideDateRange = False
            else:
                column += 1
                element = "//*[@id='root']/div/div/div/div[2]/div/div[2]/div/div[3]/a[" + str(transaction) + "]/div[" + str(column) + "]"
                description = driver.find_element(By.XPATH, element).text
                column += 2
                element = "//*[@id='root']/div/div/div/div[2]/div/div[2]/div/div[3]/a[" + str(transaction) + "]/div[" + str(column) + "]"
                amount = driver.find_element(By.XPATH, element).text.strip('+').replace('$', '').replace(',', '')
                description = modifyTransactionDescription(description, amount)
                row = m1Date, description, amount
                csv.writer(open(m1Activity, 'a', newline='')).writerow(row)
                transaction += 1
                column = 1
                element = "//*[@id='root']/div/div/div/div[2]/div/div[2]/div/div[3]/a[" + str(transaction) + "]/div[" + str(column) + "]"
        except NoSuchElementException:
            if not clickedNext:
                # click Next
                driver.find_element(By.XPATH, "//*[@id='root']/div/div/div/div[2]/div/div[2]/div/div[2]/div/button[2]/span[2]").click()
                transaction = 1
                column = 1
                element = "//*[@id='root']/div/div/div/div[2]/div/div[2]/div/div[3]/a[" + str(transaction) + "]/div[" + str(column) + "]"
                time.sleep(2)
                clickedNext = True
            else:
                insideDateRange = False


def runM1(directory, driver):
    login(directory, driver)
    m1Balance = captureBalance(driver)
    m1Activity = directory + r"\Projects\Coding\Python\BankingAutomation\Resources\m1.csv"
    gnuM1Activity = directory + r"\Projects\Coding\Python\BankingAutomation\Resources\gnu_m1.csv"
    open(m1Activity, 'w', newline='').truncate()
    open(gnuM1Activity, 'w', newline='').truncate()
    myBook = openGnuCashBook(directory, 'Finance', False, False)
    today = datetime.today()
    dateRange = getDateRange(today, 5)
    captureTransactions(driver, dateRange, m1Activity, today)
    # Compare against existing transactions in GnuCash and import new ones
    reviewTrans = compileGnuTransactions('M1', m1Activity, gnuM1Activity, myBook, driver, directory, dateRange, 0)
    return [m1Balance, reviewTrans]
    
if __name__ == '__main__':
    directory = setDirectory()
    driver = chromeDriverAsUser()
    response = runM1(directory, driver)
    print('balance: ' + response[0])
    print('transactions to review: ' + response[1])