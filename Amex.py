import os
import time
from datetime import datetime

import pyautogui
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By

from Functions import (chromeDriverAsUser, getGnuCashBalance, getPassword,
                       getUsername, importGnuTransaction, openGnuCashBook,
                       setDirectory, showMessage, updateSpreadsheet)


def login(directory, driver):
    driver.get("https://www.americanexpress.com/")
    # login
    driver.find_element(By.XPATH, "/html/body/div[1]/div/div/header/div[2]/div[1]/div[3]/div/div[5]/ul/li[3]/span/a[1]").click()
    driver.find_element(By.ID, "eliloUserID").send_keys(getUsername(directory, 'Amex'))
    driver.find_element(By.ID, "eliloPassword").send_keys(getPassword(directory, 'Amex'))
    driver.find_element(By.ID, "loginSubmit").click()
    # handle pop-up
    try:
        driver.find_element(By.XPATH, "/html/body/div[1]/div[5]/div/div/div/div/div/div[2]/div/div/div/div/div[1]/div/a/span/span").click()
    except NoSuchElementException:
        exception = "caught"
    time.sleep(1)


def captureBalance(driver):
    return driver.find_element(By.XPATH, "//*[@id='axp-balance-payment']/div[1]/div[1]/div/div[1]/div/div/span[1]/div").text.replace('$', '')


def exportTransactions(driver):
    driver.find_element(By.XPATH, "//*[@id='axp-balance-payment']/div[2]/div[2]/div/div[1]/div[1]/div/a").click()
    try: 
        # click on View Activity (for previous billing period)
        driver.find_element(By.XPATH, "//*[@id='root']/div[1]/div/div[2]/div/div/div[4]/div/div[3]/div/div/div/div/div/div/div[2]/div/div/div[5]/div/div[2]/div/div[2]/a/span").click()
    except NoSuchElementException:
        exception = "caught"
    # click on Download
    time.sleep(5)
    driver.find_element(By.XPATH, "/html/body/div[1]/div[1]/div/div[2]/div/div/div[4]/div/div[3]/div/div/div/div/div/div/div[2]/div/div/div[2]/div/div/div/div/table/thead/div/tr[1]/td[2]/div/div[2]/div/button").click()
    # click on CSV option
    driver.find_element(By.XPATH, "/html/body/div[1]/div[4]/div/div/div/div/div/div[2]/div/div[1]/div/fieldset/div[1]/label").click()
    # delete old csv file, if present
    try:
        os.remove(r"C:\Users\dmagn\Downloads\activity.csv")
    except FileNotFoundError:
        exception = "caught"
    # click on Download
    driver.find_element(By.XPATH, "/html/body/div[1]/div[4]/div/div/div/div/div/div[3]/a").click()
    time.sleep(3)

def claimRewards(driver):
    driver.find_element(By.XPATH, "/html/body/div[1]/div[1]/div/div[2]/div/div/div[4]/div/div[2]/div/div/header/div/div/div/div/div/div/div[1]/div/div[1]/div/ul/li[5]/a/span").click()
    time.sleep(10)
    rewards = driver.find_element(By.XPATH, "/html/body/div[1]/div[1]/div/div[2]/div/div[3]/div/div[3]/div/div/div[2]/div[2]/div[1]/div/div/div/div/div/div[1]/div").text.strip('$').strip(',')
    # click on Redeem for Statement Credit
    driver.find_element(By.XPATH, "/html/body/div[1]/div[1]/div/div[2]/div/div[3]/div/div[3]/div/div/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/a").click()
    try:
        # enter full rewards amount
        driver.find_element(By.XPATH, "/html/body/div[1]/div/div[4]/div/div/div[2]/section[1]/div/div[1]/p[2]/input[1]").send_keys(rewards)
        # click on Redeem Now
        driver.find_element(By.XPATH, "/html/body/div[1]/div/div[4]/div/div/div[2]/section[1]/div/div[1]/span/a").click()
        # enter email address
        driver.find_element(By.XPATH, "/html/body/div[1]/div/div[4]/div/div/div[2]/div/div/div/div/div/form/div[2]/input[2]").send_keys(os.environ.get('Email'))
        time.sleep(2)
        # accept Terms & Conditions
        driver.find_element(By.XPATH, "/html/body/div[1]/div/div[4]/div/div/div[2]/div/div/div/div/div/form/div[2]/input[2]").click()
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('space')
        # click Redeem Now (again)
        driver.find_element(By.XPATH, "/html/body/div[1]/div/div[4]/div/div/div[2]/div/div/div/div/div/form/div[2]/input[3]").click()
    except NoSuchElementException:
        exception = "caught"

def locateAndUpdateSpreadsheet(driver, amex):
    # get current date
    today = datetime.today()
    month = today.month
    year = today.year
    # switch worksheets if running in December (to next year's worksheet)
    if month == 12:
        year = year + 1
    amexNeg = float(amex.strip('$')) * -1
    updateSpreadsheet(directory, 'Checking Balance', year, 'Amex', month, amexNeg, "Amex CC")
    updateSpreadsheet(directory, 'Checking Balance', year, 'Amex', month, amexNeg, "Amex CC", True)
    # Display Checking Balance spreadsheet
    driver.execute_script("window.open('https://docs.google.com/spreadsheets/d/1684fQ-gW5A0uOf7s45p9tC4GiEE5s5_fjO5E7dgVI1s/edit#gid=1688093622');")


def runAmex(directory, driver):
    login(directory, driver)
    amex = captureBalance(driver)
    exportTransactions(driver)
    claimRewards(driver)
    myBook = openGnuCashBook(directory, 'Finance', False, False)
    reviewTrans = importGnuTransaction('Amex', r'C:\Users\dmagn\Downloads\activity.csv', myBook, driver, directory)
    amexGnu = getGnuCashBalance(myBook, 'Amex')
    locateAndUpdateSpreadsheet(driver, amex)
    if reviewTrans:
        os.startfile(directory + r"\Finances\Personal Finances\Finance.gnucash")
    showMessage("Balances + Review", f'Amex Balance: {amex} \n' f'GnuCash Amex Balance: {amexGnu} \n \n' f'Review transactions:\n{reviewTrans}')
    driver.quit()

if __name__ == '__main__':
    directory = setDirectory()
    driver = chromeDriverAsUser()
    driver.implicitly_wait(5)
    runAmex(directory, driver)
