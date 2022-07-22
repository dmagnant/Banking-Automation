import csv
import time
from datetime import datetime

from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By

from Functions import (chromeDriverAsUser, importUniqueTransactionsToGnuCash,
                       getDateRange, getPassword, getUsername,
                       modifyTransactionDescription, openGnuCashBook,
                       setDirectory, showMessage)

BALANCE_ELEMENT_XPATH = "//*[@id='root']/div/div/div/div[2]/div/div[3]/div[1]/div/div/div[1]/div/h1[2]"
                          
def login(directory, driver):
    driver.get("https://dashboard.m1finance.com/login")
    time.sleep(1)
    # login
    # enter email
    driver.find_element(By.XPATH, "/html/body/div[2]/div/div/div/div[2]/div/div/form/div[2]/div/div[1]/div/input").send_keys(getUsername(directory, 'M1 Finance'))
    # enter password
    driver.find_element(By.XPATH, "/html/body/div[2]/div/div/div/div[2]/div/div/form/div[2]/div/div[2]/div/input").send_keys(getPassword(directory, 'M1 Finance'))
    # click Log In
    driver.find_element(By.XPATH, "/html/body/div[2]/div/div/div/div[2]/div/div/form/div[4]/div/button/span[2]").click()
    time.sleep(3)
    driver.get("https://dashboard.m1.com/d/spend/checking/transactions")
    time.sleep(2)
    try:
        driver.find_element(By.XPATH, BALANCE_ELEMENT_XPATH)
    except NoSuchElementException:
        showMessage('Manual Check',"login failed or balance element not found")


def captureBalance(driver):
    return driver.find_element(By.XPATH, BALANCE_ELEMENT_XPATH).text.strip('$').replace(',', '')

def setTransactionElementRoot(row, column):
    return "//*[@id='root']/div/div/div/div[2]/div/div[3]/div[2]/div/div[3]/a[" + str(row) + "]/div[" + str(column) + "]"

def captureTransactionsFromM1Website(driver, dateRange, m1Activity, today):
    year = today.year
    row = 1
    column = 1
    elementRoot = setTransactionElementRoot(row, column)
    insideDateRange = True
    clickedNext = False

    while insideDateRange:
        try:
            # capture m1_date
            rawDate = driver.find_element(By.XPATH, elementRoot + "/span").text
            if "," in rawDate:
                try:
                    # capture Date if Mon Day, Year format (standard for prior year)
                    m1Date = datetime.strptime(driver.find_element(By.XPATH, elementRoot + "/span").text, '%b %d, %Y').date()
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
                description = driver.find_element(By.XPATH, setTransactionElementRoot(row, column) + "/div/div/h3/div").text
                column += 2
                amount = driver.find_element(By.XPATH, setTransactionElementRoot(row, column) + "/span/span").text.strip('+').replace('$', '').replace(',', '')
                description = modifyTransactionDescription(description, amount)
                transaction = m1Date, description, amount
                csv.writer(open(m1Activity, 'a', newline='')).writerow(transaction)
                row += 1
                column = 1
                elementRoot = setTransactionElementRoot(row, column)
        except NoSuchElementException:
            if not clickedNext:
                # click Next
                driver.find_element(By.XPATH, "//*[@id='root']/div/div/div/div[2]/div/div[3]/div[2]/div/div[2]/div/button[2]/span[2]").click()
                row = 1
                column = 1
                elementRoot = setTransactionElementRoot(row, column)
                time.sleep(2)
                clickedNext = True
            else:
                insideDateRange = False

def runM1(directory, driver):
    driver.implicitly_wait(5)
    login(directory, driver)
    m1Balance = captureBalance(driver)
    m1Activity = directory + r"\Projects\Coding\Python\BankingAutomation\Resources\m1.csv"
    gnuM1Activity = directory + r"\Projects\Coding\Python\BankingAutomation\Resources\gnu_m1.csv"
    open(m1Activity, 'w', newline='').truncate()
    open(gnuM1Activity, 'w', newline='').truncate()
    myBook = openGnuCashBook(directory, 'Finance', False, False)
    today = datetime.today()
    dateRange = getDateRange(today, 5)
    captureTransactionsFromM1Website(driver, dateRange, m1Activity, today)
    # Compare against existing transactions in GnuCash and import new ones
    reviewTrans = importUniqueTransactionsToGnuCash('M1', m1Activity, gnuM1Activity, myBook, driver, directory, dateRange, 0)
    return [m1Balance, reviewTrans]
    
if __name__ == '__main__':
    directory = setDirectory()
    driver = chromeDriverAsUser()
    driver.implicitly_wait(3)
    response = runM1(directory, driver)
    print('balance: ' + response[0])
    print('transactions to review: ' + response[1])
