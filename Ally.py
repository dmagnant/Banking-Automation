import csv
import time
from datetime import datetime

from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By

from Functions import (chromeDriverAsUser, closeExpressVPN,
                       importUniqueTransactionsToGnuCash, getDateRange, getPassword,
                       modifyTransactionDescription, openGnuCashBook,
                       setDirectory, showMessage)


def login(directory, driver):
    closeExpressVPN()
    driver.implicitly_wait(5)
    driver.get("https://secure.ally.com/")
    time.sleep(1)
    # login
    # enter password # changed 1/21/22
    driver.find_element(By.XPATH, "/html/body/div[1]/div[1]/main/div/div/div/div/div[1]/form/div[2]/div/span/input").send_keys(getPassword(directory, 'Ally Bank'))
    time.sleep(1)
    # click Log In
    driver.find_element(By.XPATH, "/html/body/div/div[1]/main/div/div/div/div/div[1]/form/div[3]/button/span").click()
    time.sleep(5)
    showMessage('confirm login',"manually login if necessary")

def captureBalance(driver):
    return driver.find_element(By.XPATH, "/html/body/div/div[1]/main/div/div/div/div[2]/div/div[2]/div/table/tbody/tr/td[2]/div").text.replace('$', '').replace(',', '')

def captureTransactions(driver, dateRange, allyActivity):
     # click Joint Checking link
    driver.find_element(By.PARTIAL_LINK_TEXT, "Joint Checking").click()
    time.sleep(3)
    table = 2
    transaction = 1
    column = 1
    element = "//*[@id='form-elements-section']/section/section/table[" + str(table) + "]/tbody/tr[" + str(transaction) + "]/td[" + str(column) + "]"
    time.sleep(6)
    insideDateRange = True
    while insideDateRange:
        try:
            modDate = datetime.strptime(driver.find_element(By.XPATH, element).text, '%b %d, %Y').date()
            if str(modDate) not in dateRange:
                insideDateRange = False
            else:
                column += 1
                element = "//*[@id='form-elements-section']/section/section/table[" + str(table) + "]/tbody/tr[" + str(transaction) + "]/td[" + str(column) + "]/button"
                description = driver.find_element(By.XPATH, element).text
                column += 1
                element = "//*[@id='form-elements-section']/section/section/table[" + str(table) + "]/tbody/tr[" + str(transaction) + "]/td[" + str(column) + "]"
                amount = driver.find_element(By.XPATH, element).text.replace('$','').replace(',','')
                description = modifyTransactionDescription(description)
                row = str(modDate), description, amount
                csv.writer(open(allyActivity, 'a', newline='')).writerow(row)
                transaction += 2
                column = 1
                element = "//*[@id='form-elements-section']/section/section/table[" + str(table) + "]/tbody/tr[" + str(transaction) + "]/td[" + str(column) + "]"
        except NoSuchElementException:
            table += 1
            column = 1
            transaction = 1
            element = "//*[@id='form-elements-section']/section/section/table[" + str(table) + "]/tbody/tr[" + str(transaction) + "]/td[" + str(column) + "]"
        except ValueError:
            insideDateRange = False

def runAlly(directory, driver):
    login(directory, driver)
    ally = captureBalance(driver)
    allyActivity = directory + r"\Projects\Coding\Python\BankingAutomation\Resources\ally.csv"
    open(allyActivity, 'w', newline='').truncate()
    # today = datetime.today()
    dateRange = getDateRange(datetime.today(), 8) # set range of days to look for transactions
    captureTransactions(driver, dateRange, allyActivity)
    myBook = openGnuCashBook(directory, 'Home', False, False)
    gnuAllyActivity = directory + r"\Projects\Coding\Python\BankingAutomation\Resources\gnu_ally.csv"
    open(gnuAllyActivity, 'w', newline='').truncate()
    # Compare against existing transactions in GnuCash and import new ones
    reviewTrans = importUniqueTransactionsToGnuCash('Ally', allyActivity, gnuAllyActivity, myBook, driver, directory, dateRange, 0)
    return [ally, reviewTrans]

if __name__ == '__main__':
    directory = setDirectory()
    driver = chromeDriverAsUser()
    response = runAlly(directory, driver)
    print('balance: ' + response[0])
    print('transactions to review: ' + response[1])
