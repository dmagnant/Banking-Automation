from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from decimal import Decimal
from datetime import datetime
import time
import os
import pyautogui
from Functions import setDirectory, chromeDriverAsUser, getUsername, getPassword, openGnuCashBook, showMessage, getGnuCashBalance, updateSpreadsheet, getStartAndEndOfPreviousMonth, writeGnuTransaction 

def login(directory, driver):
    driver.get("https://ownyourfuture.vanguard.com/login#/")
    # Enter username
    driver.find_element(By.ID, "username").send_keys(getUsername(directory, 'Vanguard'))
    time.sleep(1)
    # Enter password
    driver.find_element(By.ID, "pword").send_keys(getPassword(directory, 'Vanguard'))
    time.sleep(1)
    # click Log in
    driver.find_element(By.XPATH, "//*[@id='vui-button-1']/button/div").click()
    # handle security code
    try:
        driver.find_element(By.ID, 'vui-radio-1-input-label').click()
        showMessage('Security Code', "Enter Security code, then click OK")
        driver.find_element(By.XPATH, "//*[@id='security-code-submit-btn']/button/div").click()
    except NoSuchElementException:
        exception = "caught"


def captureBalances(driver):
    # navigate to asset details page (click view all assets)
    driver.get('https://ownyourfuture.vanguard.com/main/dashboard/assets-details')
    time.sleep(2)
    # move cursor to middle window
    pyautogui.moveTo(500, 500)
    #scroll down
    pyautogui.scroll(-1000)
    # Get Total Account Balance
    vanguard = driver.find_element(By.XPATH, "/html/body/div[3]/div/app-dashboard-root/app-assets-details/app-balance-details/div/div[3]/div[3]/div/app-details-card/div/div/div[1]/div[3]/h4").text.strip('$').replace(',', '')
    # Get Interest YTD
    interestYTD = driver.find_element(By.XPATH, "/html/body/div[3]/div/app-dashboard-root/app-assets-details/app-balance-details/div/div[3]/div[4]/div/app-details-card/div/div/div[1]/div[3]/h4").text.strip('$').replace(',', '')
    return [vanguard, interestYTD]


def importGnuTransactions(myBook, today, balances):
    #get current date
    year = today.year
    month = today.month
    lastMonth = getStartAndEndOfPreviousMonth(today, month, year)
    interestAmount = 0
    pension = getGnuCashBalance(myBook, 'VanguardPension')
    pensionAcct = "Assets:Non-Liquid Assets:Pension"
    with myBook as book:
        # # GNUCASH
        # retrieve transactions from GnuCash
        transactions = [tr for tr in book.transactions
                        if str(tr.post_date.strftime('%Y')) == str(lastMonth[0].year)
                        for spl in tr.splits
                        if spl.account.fullname == pensionAcct
                        ]
        for tr in transactions:
            date = str(tr.post_date.strftime('%Y'))
            for spl in tr.splits:
                if spl.account.fullname == "Income:Investments:Interest":
                    interestAmount = interestAmount + abs(spl.value)
        accountChange = Decimal(balances[0]) - pension
        interest = Decimal(balances[1]) - interestAmount
        employerContribution = accountChange - interest
        writeGnuTransaction(myBook, "Contribution + Interest", lastMonth[1], [-interest, -employerContribution, accountChange], pensionAcct)   
    book.close()

    return [interest, employerContribution]

def runVanguard(directory, driver): 
    login(directory, driver)
    balances = captureBalances(driver)
    myBook = openGnuCashBook(directory, 'Finance', False, False)
    today = datetime.today()
    values = importGnuTransactions(myBook, today, balances)
    vanguardGnu = getGnuCashBalance(myBook, 'VanguardPension')
    updateSpreadsheet(directory, 'Asset Allocation', today.year, 'VanguardPension', today.month, float(balances[0]))
    os.startfile(directory + r"\Finances\Personal Finances\Finance.gnucash")
    driver.execute_script("window.open('https://docs.google.com/spreadsheets/d/1sWJuxtYI-fJ6bUHBWHZTQwcggd30RcOSTMlqIzd1BBo/edit#gid=2058576150');")
    showMessage("Balances",f'Pension Balance: {balances[0]} \n'f'GnuCash Pension Balance: {vanguardGnu} \n'f'Interest earned: {values[0]} \n'f'Total monthly contributions: {values[1]} \n')
    driver.quit()

if __name__ == '__main__':
    directory = setDirectory()
    driver = chromeDriverAsUser()
    driver.implicitly_wait(5)
    runVanguard(directory, driver)
    