import time
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, ElementNotInteractableException
import os
from datetime import datetime
from Functions import setDirectory, chromeDriverAsUser, chromeDriverBlank, getUsername, getPassword, openGnuCashBook, showMessage, importGnuTransaction, getGnuCashBalance, updateSpreadsheet


def login(directory, driver):
    driver.get("https://portal.discover.com/customersvcs/universalLogin/ac_main")
    # login
    driver.find_element(By.ID, 'userid-content').send_keys(getUsername(directory, 'Discover'))
    driver.find_element(By.ID, 'password-content').send_keys(getPassword(directory, 'Discover'))
    time.sleep(1)
    driver.find_element(By.XPATH, '/html/body/div[1]/main/div/div[1]/div/form/input[8]').click()

    #handle pop-up
    try:
        driver.find_element(By.XPATH, "/html/body/div[1]/main/div[12]/div/div/div[2]/a").click()
    except (NoSuchElementException, ElementNotInteractableException, AttributeError):
        exception = "caught"
    showMessage("Login Check", 'Confirm Login to , (manually if necessary) \n' 'Then click OK \n')


def captureBalance(driver):
    return driver.find_element(By.XPATH, "/html/body/div[1]/main/div[4]/div/div[1]/div[1]/p[3]/span[2]").text


def exportTransactions(driver, today):
    # Click on All Activity & Statements
    driver.find_element(By.PARTIAL_LINK_TEXT, "All Activity & Statements").click()
    # Click on "Select Activity or Statement Period"
    driver.find_element(By.XPATH, "/html/body/div[1]/main/div[1]/div[4]/div/div[1]/div/div/div/div/div[1]/a").click()
    # Click on Current
    driver.find_element(By.PARTIAL_LINK_TEXT, "Current").click()
    driver.implicitly_wait(3)
    # Click on Download
    driver.find_element(By.XPATH, "/html/body/div[1]/main/div[1]/aside/div[2]/div[1]/div/a[2]").click()
    # CLick on CSV
    driver.find_element(By.ID, "radio4").click()
    # CLick Download
    driver.find_element(By.ID, "submitDownload").click()
    # Click Close
    driver.find_element(By.XPATH, "/html/body/div[1]/main/div[5]/div/form/div/div[4]/a[1]").click()
    year = today.year
    stmtYear = str(year)
    stmtMonth = today.strftime('%m')
    return r"C:\Users\dmagn\Downloads\Discover-Statement-" + stmtYear + stmtMonth + "12.csv"


def claimRewards(driver):
    # Click Rewards
    driver.find_element(By.XPATH, "/html/body/div[1]/header/div/div/div[3]/div[1]/ul/li[4]/a").click()
    # Click "Redeem Cashback Bonus"
    driver.find_element(By.XPATH, "/html/body/div[1]/header/div/div/div[3]/div[1]/ul/li[4]/div/div/div/ul/li[3]/a").click()
    # Click "Cash It"
    driver.find_element(By.XPATH, "//*[@id='redemption-module']/li[1]/a").click()
    time.sleep(1)
    try:
        # Click Electronic Deposit to your bank account
        driver.find_element(By.XPATH, "//*[@id='electronic-deposit']").click()
        time.sleep(1)
        # Click Redeem All link
        driver.find_element(By.XPATH, "/html/body/div[1]/div[1]/main/div/div/section/div[2]/div/form/div[2]/fieldset/div[3]/div[2]/span[2]/button").click()
        time.sleep(1)
        # Click Continue
        driver.find_element(By.XPATH, "/html/body/div[1]/div[1]/main/div/div/section/div[2]/div/form/div[4]/input").click()
        time.sleep(1)
        # Click Submit
        driver.find_element(By.XPATH, "/html/body/div[1]/div[1]/main/div/div/section/div[2]/div/div/div/div[1]/div/div/div[2]/div/button[1]").click()
    except NoSuchElementException:
        exception = "caught"


def locateAndUpdateSpreadsheet(driver, discover, today):
    # switch worksheets if running in December (to next year's worksheet)
    month = today.month
    year = today.year
    if month == 12:
        year = year + 1
    discoverNeg = float(discover.strip('$')) * -1
    updateSpreadsheet(directory, 'Checking Balance', year, 'Discover', month, discoverNeg, 'Discover CC')
    updateSpreadsheet(directory, 'Checking Balance', year, 'Discover', month, discoverNeg, 'Discover CC', True)

    # Display Checking Balance spreadsheet
    driver.execute_script("window.open('hhttps://docs.google.com/spreadsheets/d/1684fQ-gW5A0uOf7s45p9tC4GiEE5s5_fjO5E7dgVI1s/edit#gid=1688093622');")


def runDiscover(directory, driver):
    login(directory, driver)
    discover = captureBalance(driver)
    today = datetime.today()
    transactionsCSV = exportTransactions(driver, today)
    claimRewards(driver)
    myBook = openGnuCashBook(directory, 'Finance', False, False)
    reviewTrans = importGnuTransaction('Discover', transactionsCSV, myBook, driver, directory)
    discoverGnu = getGnuCashBalance(myBook, 'Discover')
    locateAndUpdateSpreadsheet(driver, discover, today)
    if reviewTrans:
        os.startfile(directory + r"\Finances\Personal Finances\Finance.gnucash")
    showMessage("Balances + Review", f'Discover Balance: {discover} \n' f'GnuCash Discover Balance: {discoverGnu} \n \n' f'Review transactions:\n{reviewTrans}')
    driver.quit()

if __name__ == '__main__':
    directory = setDirectory()
    driver = chromeDriverAsUser()
    driver.implicitly_wait(6)
    runDiscover(directory, driver)
