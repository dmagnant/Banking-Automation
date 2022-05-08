from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, ElementNotInteractableException
from selenium.webdriver.common.keys import Keys
from datetime import datetime
import time
import os
from Functions import setDirectory, chromeDriverAsUser, getUsername, getPassword, openGnuCashBook, showMessage, getGnuCashBalance, updateSpreadsheet, importGnuTransaction, startExpressVPN

def runBoA(directory, driver, account):
    # account = p for personal or j for joint
    driver.implicitly_wait(3)
    driver.get("https://www.bankofamerica.com/")
    # login
    driver.find_element(By.ID, "onlineId1").send_keys(getUsername(directory, 'BoA CC'))
    driver.find_element(By.ID, "passcode1").send_keys(getPassword(directory, 'BoA CC'))
    driver.find_element(By.XPATH, "//*[@id='signIn']").click()
    # handle ID verification
    try:
        driver.find_element(By.XPATH, "//*[@id='btnARContinue']/span[1]").click()
        showMessage("Get Verification Code", "Enter code, then click OK")
        driver.find_element(By.XPATH, "//*[@id='yes-recognize']").click()
        driver.find_element(By.XPATH, "//*[@id='continue-auth-number']/span").click()
    except NoSuchElementException:
        exception = "Caught"

    # handle security questions
    try:
        question = driver.find_element(By.XPATH, "/html/body/div[1]/div/div/div[2]/div[1]/div/div/form/div[2]/label").text
        q1 = "What is the name of a college you applied to but didn't attend?"
        a1 = os.environ.get('CollegeApplied')
        q2 = "As a child, what did you want to be when you grew up?"
        a2 = os.environ.get('DreamJob')
        q3 = "What is the name of your first babysitter?"
        a3 = os.environ.get('FirstBabySitter')
        if driver.find_element(By.XPATH, "/html/body/div[1]/div/div/div[2]/div[1]/div/div/form/div[2]/label"):
            question = driver.find_element(By.XPATH, 
                "/html/body/div[1]/div/div/div[2]/div[1]/div/div/form/div[2]/label").text
            if question == q1:
                driver.find_element(By.NAME, "challengeQuestionAnswer").send_keys(a1)
            elif question == q2:
                driver.find_element(By.NAME, "challengeQuestionAnswer").send_keys(a2)
            else:
                driver.find_element(By.NAME, "challengeQuestionAnswer").send_keys(a3)
            driver.find_element(By.XPATH, "/html/body/div[1]/div/div/div[2]/div[1]/div/div/form/fieldset/div[2]/div/div[1]/input").click()
            driver.find_element(By.XPATH, "/html/body/div[1]/div/div/div[2]/div[1]/div/div/form/a[1]/span").click()
    except NoSuchElementException:
        exception = "Caught"

    # close mobile app pop-up
    try:
        driver.find_element(By.XPATH, "//*[@id='sasi-overlay-module-modalClose']/span[1]").click()
    except NoSuchElementException:
        exception = "Caught"

    # account = p for personal or j for joint
    partial_link = 'Customized Cash Rewards Visa Signature - 8549' if account == 'p' else 'Travel Rewards Visa Signature - 8955'

    driver.find_element(By.PARTIAL_LINK_TEXT, partial_link).click()
    # # Capture Statement balance
    BoA = driver.find_element(By.XPATH, "/html/body/div[1]/div/div[2]/div/div[2]/div[2]/div/div/div[3]/div[4]/div[3]/div/div[2]/div[2]/div[2]").text.replace('$','').replace(',','')
    # # EXPORT TRANSACTIONS
    # click Previous transactions
    time.sleep(3)
    driver.find_element(By.PARTIAL_LINK_TEXT, "Previous transactions").click()
    # click Download, select microsoft excel
    ## had to edit div1/div2 on 1/19/22
    try: 
        driver.find_element(By.XPATH, "/html/body/div[1]/div/div[4]/div[1]/div/div[5]/div[2]/div[1]/div/div[1]/a").click()
        driver.find_element(By.XPATH, "/html/body/div[1]/div/div[4]/div[1]/div/div[5]/div[2]/div[1]/div/div[3]/div/div[3]/div[1]/select").send_keys("m")
    except NoSuchElementException:
        driver.find_element(By.XPATH, "/html/body/div[1]/div/div[4]/div[1]/div/div[5]/div[2]/div[2]/div/div[1]/a").click()
        driver.find_element(By.XPATH, "/html/body/div[1]/div/div[4]/div[1]/div/div[5]/div[2]/div[2]/div/div[3]/div/div[3]/div[1]/select").send_keys("m")
    driver.execute_script("window.scrollTo(0, 300)")
    # click Download Transactions
    try: 
        driver.find_element(By.XPATH, "/html/body/div[1]/div/div[4]/div[1]/div/div[5]/div[2]/div[1]/div/div[3]/div/div[4]/div[2]/a/span").click()
    except NoSuchElementException:
        driver.find_element(By.XPATH, "/html/body/div[1]/div/div[4]/div[1]/div/div[5]/div[2]/div[2]/div/div[3]/div/div[4]/div[2]/a/span").click()
    # get current date
    today = datetime.today()
    year = today.year
    month = today.month
    stmtmonth = today.strftime("%B")
    stmtyear = str(year)
    account_num = "_8549.csv" if account == 'p' else "_8955.csv"
    transactions_csv = os.path.join(r"C:\Users\dmagn\Downloads", stmtmonth + stmtyear + account_num)
    time.sleep(2)
    # Set Gnucash Book
    mybook = openGnuCashBook(directory, 'Finance', False, False) if account == 'p' else openGnuCashBook(directory, 'Home', False, False)

    import_acct = 'BoA' if account == 'p' else 'BoA-joint'
    review_trans = importGnuTransaction(import_acct, transactions_csv, mybook, driver, directory)
    BoA_gnu = getGnuCashBalance(mybook, import_acct)

    if account == 'p':
        # # REDEEM REWARDS
        # click on View/Redeem menu
        driver.find_element(By.XPATH, "/html/body/div[1]/div/div[2]/div/div[2]/div[2]/div/div/div[1]/div[4]/div[3]/a").click()
        time.sleep(5)
        #scroll down to view button
        driver.execute_script("window.scrollTo(0, 300)")
        time.sleep(2)
        # wait for Redeem Cash Rewards button to load, click it
        driver.find_element(By.ID, "rewardsRedeembtn").click()
        # switch to last window
        driver.switch_to.window(driver.window_handles[len(driver.window_handles)-1])
        # close out of pop-up (if present)
        try:
            driver.find_element(By.XPATH, "/html/body/div[1]/div/div/div[2]/div/div/button").click()
        except NoSuchElementException:
            exception = "caught"
        # if no pop-up, proceed to click on Redemption option
        driver.find_element(By.ID, "redemption_option").click()
        # redeem if there is a balance, else skip
        try:
            # Choose Visa - statement credit
            driver.find_element(By.ID, "redemption_option").send_keys("v")
            driver.find_element(By.ID, "redemption_option").send_keys(Keys.ENTER)
            # click on Redeem all
            driver.find_element(By.XPATH, "/html/body/main/div[2]/div/div/div/div/div[1]/div[2]/div[1]/div/div/form/button").click()
            # click Complete Redemption
            driver.find_element(By.XPATH, "/html/body/div[1]/div/div/div[3]/button[1]").click()
        except ElementNotInteractableException:
            exception = "caught"

    # switch worksheets if running in December (to next year's worksheet)
    if month == 12:
        year = year + 1
    BoA_neg = float(BoA) * -1

    if account == 'p':
        updateSpreadsheet(directory, 'Checking Balance', year, 'BoA', month, BoA_neg, 'BoA CC')
        updateSpreadsheet(directory, 'Checking Balance', year, 'BoA', month, BoA_neg, 'BoA CC', True)
        driver.execute_script("window.open('https://docs.google.com/spreadsheets/d/1684fQ-gW5A0uOf7s45p9tC4GiEE5s5_fjO5E7dgVI1s/edit#gid=1688093622');")
    else:
        updateSpreadsheet(directory, 'Home', str(year) + ' Balance', 'BoA-joint', month, BoA_neg, 'BoA CC')
        updateSpreadsheet(directory, 'Home', str(year) + ' Balance', 'BoA-joint', month, BoA_neg, 'BoA CC', True)
        # Display Home spreadsheet
        driver.execute_script("window.open('https://docs.google.com/spreadsheets/d/1oP3U7y8qywvXG9U_zYXgjFfqHrCyPtUDl4zPDftFCdM/edit#gid=317262693');")

    # Display Checking Balance spreadsheet
    # Open GnuCash if there are transactions to review
    if review_trans:
        os.startfile(directory + r"\Finances\Personal Finances\Finance.gnucash") if account == 'p' else os.startfile(directory + r"\Stuff\Home\Finances\Home.gnucash")
    # Display Balance
    showMessage("Balances + Review", f'BoA Balance: {BoA} \n' f'GnuCash BoA Balance: {BoA_gnu} \n \n' f'Review transactions:\n{review_trans}')
    driver.quit()
    # startExpressVPN()

if __name__ == '__main__':
    directory = setDirectory()
    driver = chromeDriverAsUser()
    account = 'j'
    runBoA(directory, driver, account)