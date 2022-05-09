from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, ElementClickInterceptedException
import time
from datetime import datetime
import os
from Functions import setDirectory, chromeDriverAsUser, getUsername, getPassword, openGnuCashBook, showMessage, getGnuCashBalance, updateSpreadsheet, importGnuTransaction

def login(driver):
    driver.get("https://www.chase.com/")
    time.sleep(2)
    # login
    showMessage("Login Manually", 'login manually \n' 'Then click OK \n')
    driver.get("https://secure07a.chase.com/web/auth/dashboard#/dashboard/overviewAccounts/overview/multiProduct;flyout=accountSummary,818208017,CARD,BAC")
    # time.sleep(2)



def captureBalance(driver):
    return driver.find_element(By.XPATH, "/html/body/div[2]/div/div[1]/div[2]/main/div[3]/div/div/div/div/div/div/div[1]/section[1]/div[6]/div/div[2]/div/div/div[3]/div/div[2]/ul/li[2]/div/div[2]/span").text.strip('$')


def exportTransactions(driver, today):
    # click on Activity since last statement
    driver.find_element(By.XPATH, "//*[@id='header-transactionTypeOptions']/span[1]").click()
    # choose last statement
    driver.find_element(By.ID, "item-0-STMT_CYCLE_1").click()
    time.sleep(1)
    # click Download
    driver.find_element(By.XPATH, "//*[@id='downloadActivityIcon']").click()
    time.sleep(2)
    # click Download
    driver.find_element(By.ID, "download").click()

    year = today.year
    month = today.month
    day = today.strftime('%d')
    monthTo = today.strftime('%m')
    yearTo = str(year)
    monthFrom = "12"         if month == 1 else "{:02d}".format(month - 1)
    yearfrom = str(year - 1) if month == 1 else yearTo
    fromDate = yearfrom + monthFrom + "07_"
    toDate = yearTo + monthTo + "06_"
    currentDate = yearTo + monthTo + day
    return r'C:\Users\dmagn\Downloads\Chase2715_Activity' + fromDate + toDate + currentDate + '.csv'


def claimRewards(driver):
    driver.get("https://ultimaterewardspoints.chase.com/cash-back?lang=en")
    try:
        # Deposit into a Bank Account
        driver.find_element(By.XPATH, "/html/body/the-app/main/ng-component/main/div/section[2]/div[2]/form/div[6]/ul/li[2]/label").click()
        # Click Continue
        driver.find_element(By.XPATH, "/html/body/the-app/main/ng-component/main/div/section[2]/div[2]/form/div[7]/button").click()
        # Click Confirm & Submit
        driver.find_element(By.ID, "cash_back_button_submit").click()
    except NoSuchElementException:
        exception = "caught"
    except ElementClickInterceptedException:
        exception = "caught"


def locateAndUpdateSpreadsheet(driver, chase, today):
    # switch worksheets if running in December (to next year's worksheet)
    month = today.month
    year = today.year
    if month == 12:
        year = year + 1
    chaseNeg = float(chase) * -1
    updateSpreadsheet(directory, 'Checking Balance', year, 'Chase', month, chaseNeg, 'Chase CC')
    updateSpreadsheet(directory, 'Checking Balance', year, 'Chase', month, chaseNeg, 'Chase CC', True)
    # Display Checking Balance spreadsheet
    driver.execute_script("window.open('https://docs.google.com/spreadsheets/d/1684fQ-gW5A0uOf7s45p9tC4GiEE5s5_fjO5E7dgVI1s/edit#gid=1688093622');")


def runChase(directory, driver):
    login(driver)
    chase = captureBalance(driver)
    today = datetime.today()
    transactionsCSV = exportTransactions(driver, today)
    claimRewards(driver)
    myBook = openGnuCashBook(directory, 'Finance', False, False)
    reviewTrans = importGnuTransaction('Chase', transactionsCSV, myBook, driver, directory)
    chaseGnu = getGnuCashBalance(myBook, 'Chase')
    locateAndUpdateSpreadsheet(driver, chase, today)
    if reviewTrans:
        os.startfile(directory + r"\Finances\Personal Finances\Finance.gnucash")
    showMessage("Balances + Review", f'Chase Balance: {chase} \n' f'GnuCash Chase Balance: {chaseGnu} \n \n' f'Review transactions:\n{reviewTrans}')
    driver.quit()

if __name__ == '__main__':
    directory = setDirectory()
    driver = chromeDriverAsUser()
    driver.implicitly_wait(5)
    runChase(directory, driver)
