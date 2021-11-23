from ahk import AHK
import ctypes
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
from datetime import datetime
import time
from decimal import Decimal
from piecash import Transaction, Split
import pyotp
import pyautogui
from Functions import setDirectory, chromeDriverAsUser, getUsername, getPassword, openGnuCashBook, showMessage, getGnuCashBalance, updateSpreadsheet, getStartAndEndOfPreviousMonth, writeGnuTransaction


directory = setDirectory()
driver = chromeDriverAsUser(directory)
driver.implicitly_wait(6)
driver.get("https://www.myconstant.com/log-in")
ahk = AHK()
#login
try:
    driver.find_element_by_id("lg_username").send_keys(getUsername(directory, 'My Constant'))
    driver.find_element_by_id("lg_password").send_keys(getPassword(directory, 'My Constant'))
    driver.find_element_by_id("lg_password").click()
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('space')
    # ahk.key_press('Tab')
    # ahk.key_press('Tab')
    # ahk.key_press('Space')
    # handle captcha
    MessageBox = ctypes.windll.user32.MessageBoxW
    MessageBox(None, "Verify captcha, then click OK", 'CAPTCHA', 0)
    driver.maximize_window()
    driver.find_element_by_xpath("//*[@id='submit-btn']").click()
    totp = pyotp.TOTP("IM26VSEB6FOMXVHKO6Z7XSGRVL5HUNCI")
    token = totp.now()
    char = 0
    time.sleep(2)
    while char < 6:
        xpath_start = "//*[@id='layout']/div[2]/div/div/div[2]/div/div/div/div[3]/div/div/div["
        driver.find_element_by_xpath(xpath_start + str(char + 1) + "]/input").send_keys(token[char])
        char += 1
    time.sleep(6)
except NoSuchElementException:
    exception = "caught"
pyautogui.moveTo(1650, 175)
time.sleep(8)
# capture and format Constant balance
constant_balance_raw = driver.find_element_by_id("acc_balance").text.replace(',', '').replace('$', '')
constant_balance_dec = Decimal(constant_balance_raw)
constant_balance_rnd = str(round(constant_balance_dec, 2))
constant_balance = float(constant_balance_rnd)

# # # WORTHY
# # LOAD PAGE
driver.execute_script("window.open('https://worthy.capital/start');")
# make the new window active
window_after = driver.window_handles[1]
driver.switch_to.window(window_after)
time.sleep(2)
try:
    # login
    # click Login button
    driver.find_element_by_xpath("//*[@id='q-app']/div/div[1]/div/div[2]/div/button[2]/span[2]/span").click()
    # enter credentials
    driver.find_element_by_xpath("//*[@id='auth0-lock-container-1']/div/div[2]/form/div/div/div[2]/span/div/div/div/div/div/div/div/div/div[2]/div[1]/div[1]/input").send_keys(getUsername(directory, 'Worthy'))
    driver.find_element_by_xpath("//*[@id='auth0-lock-container-1']/div/div[2]/form/div/div/div[2]/span/div/div/div/div/div/div/div/div/div[2]/div[2]/div/div[1]/input").send_keys(getPassword(directory, 'Worthy'))
    # click Login button (again)
    driver.find_element_by_xpath("//*[@id='auth0-lock-container-1']/div/div[2]/form/div/div/button/span").click()
except NoSuchElementException:
    exception = "caught"
# capture balances
time.sleep(3)
# Get balance from Worthy I

driver.find_element_by_xpath("//*[@id='q-app']/div/div[1]/main/div/div/div[2]/div/div[2]/div/div/div[1]").click()
wpc1_raw = driver.find_element_by_xpath("/html/body/div[1]/div/div[1]/main/div/div/div[2]/div/div[2]/div/div/div[3]/div/h4/span[3]").text.replace('$', '').replace(',', '')
wpc1 = Decimal(wpc1_raw)
# Get balance from Worthy II
driver.find_element_by_xpath("/html/body/div[1]/div/div[1]/main/div/div/div[2]/div/div[3]/div/div/div[2]/div/h4/span[3]").click()
wpc2_raw = driver.find_element_by_xpath("/html/body/div[1]/div/div[1]/main/div/div/div[2]/div/div[3]/div/div/div[2]/div/h4/span[3]").text.replace('$', '').replace(',', '')
wpc2 = Decimal(wpc2_raw)
# Combine Worthy balances
worthy_balance_dec = wpc1 + wpc2
# convert from decimal to float
worthy_balance = float(worthy_balance_dec)
# calculate total bond balance
bond_balance = worthy_balance + constant_balance
# Set Gnucash Book
mybook = openGnuCashBook(directory, 'Finance', False, False)
# Obtain balances from Gnucash
constant_old_bal = float(mybook.accounts(fullname="Assets:Liquid Assets:My Constant").get_balance())
worthy_old_bal = float(mybook.accounts(fullname="Assets:Liquid Assets:Worthy Bonds").get_balance())
# Calculate Interest
constant_interest = Decimal(constant_balance - constant_old_bal)
worthy_interest = Decimal(worthy_balance - worthy_old_bal)
# Get current date
today = datetime.today()
year = today.year
month = today.month

lastmonth = getStartAndEndOfPreviousMonth(today, month, year)

with mybook as book:
    # Add Interest Transaction to My Constant
    writeGnuTransaction(book, "Interest", lastmonth[1].date(), -round(constant_interest, 2), "Income:Investments:Interest", "Assets:Liquid Assets:My Constant", '')
    # Add Interest Transaction to Worthy Bonds
    writeGnuTransaction(book, "Interest", lastmonth[1].date(), -round(worthy_interest, 2), "Income:Investments:Interest", "Assets:Liquid Assets:Worthy Bonds", '')

# # Add Interest Transactions
# with mybook as book:
#     # Add Interest Transaction to My Constant
#     to_account = "Assets:Liquid Assets:My Constant"
#     amount = round(constant_interest, 2)
#     from_account = "Income:Investments:Interest"
#     trans1 = Transaction(post_date=lastmonth[1].date(),
#                          currency=mybook.currencies(mnemonic="USD"),
#                          description="Interest",
#                          splits=[
#                             Split(value=amount, account=mybook.accounts(fullname=to_account)),
#                             Split(value=-amount, account=mybook.accounts(fullname=from_account)),
#                         ])
#     # Add Interest Transaction to Worthy Bonds
#     to_account = "Assets:Liquid Assets:Worthy Bonds"
#     amount = round(worthy_interest, 2)
#     from_account = "Income:Investments:Interest"
#     trans2 = Transaction(post_date=lastmonth[1].date(),
#                          currency=mybook.currencies(mnemonic="USD"),
#                          description="Interest",
#                          splits=[
#                              Split(value=amount, account=mybook.accounts(fullname=to_account)),
#                              Split(value=-amount, account=mybook.accounts(fullname=from_account)),
#                          ])
#     book.save()
#     book.flush()
# book.close()
# get Liquid Asset Balance from GnuCash
liq_assets = getGnuCashBalance(mybook, 'Liquid Assets')

# # # Health Equity HSA
driver.execute_script("window.open('https://member.my.healthequity.com/hsa/21895515-010');")
# make the new window active
window_after = driver.window_handles[2]
driver.switch_to.window(window_after)
# Login
# click Login
driver.find_element_by_id("ctl00_modulePageContent_btnLogin").click()
# Two-Step Authentication
try:
    # send code to phone
    driver.find_element_by_xpath("//*[@id='sendEmailTextVoicePanel']/div[5]/span[1]/span/label/span/strong").click()
    # Send confirmation code
    driver.find_element_by_id("sendOtp").click()
    # enter text code
    showMessage("Confirmation Code", "Enter code then click OK")
    # Remember me
    driver.find_element_by_xpath("//*[@id='VerifyOtpPanel']/div[4]/div[1]/div/label/span").click()
    # click Confirm
    driver.find_element_by_id("verifyOtp").click()
except NoSuchElementException:
    exception = "already verified"
# Capture balances
HE_hsa_avail_bal = driver.find_element_by_xpath("//*[@id='21895515-020']/div/hqy-hsa-tab/div/div[2]/div/span[1]").text.replace('$', '').replace(',', '')
HE_hsa_invest_bal = driver.find_element_by_xpath("//*[@id='21895515-020']/div/hqy-hsa-tab/div/div[2]/span[2]/span[1]").text.replace('$', '').replace(',', '')
HE_hsa_invest_total = float(HE_hsa_avail_bal) + float(HE_hsa_invest_bal)
HE_hsa_balance = float(HE_hsa_invest_total)
vanguard401k = driver.find_element_by_xpath("//*[@id='retirementAccounts']/li/a/div/ul/li/span[2]").text.replace('$', '').replace(',', '')
vanguard401kbal = float(vanguard401k)
# click Manage HSA Investments
driver.find_element_by_xpath("//*[@id='hsaInvestment']/div/div/a").click()
time.sleep(1)
# click Portfolio performance
driver.find_element_by_id("EditPortfolioTab").click()
time.sleep(4)
# enter Start Date
num = 0
while num < 10:
    driver.find_element_by_id("startDate").click()
    driver.find_element_by_id("startDate").send_keys(Keys.BACKSPACE)
    num += 1
driver.find_element_by_id("startDate").send_keys(datetime.strftime(lastmonth[0], '%m/%d/%Y'))
# enter End Date
num = 0
while num < 10:
    driver.find_element_by_id("endDate").click()
    driver.find_element_by_id("endDate").send_keys(Keys.BACKSPACE)
    num += 1
driver.find_element_by_id("endDate").send_keys(datetime.strftime(lastmonth[1], '%m/%d/%Y'))
# click Refresh
driver.find_element_by_id("fundPerformanceRefresh").click()
# Capture Dividends
HE_hsa_dividends = Decimal(driver.find_element_by_xpath("//*[@id='EditPortfolioTab-panel']/member-portfolio-edit-display/member-overall-portfolio-performance-display/div[1]/div/div[3]/div/span").text.replace('$', '').replace(',', ''))

# capture GnuCash HSA Balance
HSA_gnu_balance = getGnuCashBalance(mybook, 'Liquid Assets')

# Capture HSA change
HE_hsa_change = Decimal(HE_hsa_balance - HSA_gnu_balance)
# Capture market change
HE_hsa_mkt_change = Decimal(HE_hsa_change - HE_hsa_dividends)
# codifying data
HE_hsa_change = round(HE_hsa_change, 2)
HE_hsa_mkt_change = round(HE_hsa_mkt_change, 2)

# Add HSA Transactions
with mybook as book:
    writeGnuTransaction(mybook, "HSA Statement", lastmonth[1].date(), [HE_hsa_change, -HE_hsa_dividends, -HE_hsa_mkt_change], ["Income:Investments:Dividends", "Income:Investments:Market Change"], "Assets:Non-Liquid Assets:HSA", '')

    # USD = mybook.currencies(mnemonic="USD")
    # # Add Interest Transaction to My Constant
    # to_account = "Assets:Non-Liquid Assets:HSA"
    # trans1 = Transaction(post_date=lastmonth[1].date(),
    #                      currency=USD,
    #                      description="HSA Statement",
    #                      splits=[
    #                         Split(value=HE_hsa_change, account=mybook.accounts(fullname=to_account)),
    #                         Split(value=-HE_hsa_dividends, account=mybook.accounts(fullname="Income:Investments:Dividends")),
    #                         Split(value=-HE_hsa_mkt_change, account=mybook.accounts(fullname="Income:Investments:Market Change")),
    #                     ])
    # book.save()
    # book.flush()
book.close()
updateSpreadsheet(directory, 'Asset Allocation', year, 'Bonds', month, bond_balance)
updateSpreadsheet(directory, 'Asset Allocation', year, 'HE_HSA', month, HE_hsa_balance)
updateSpreadsheet(directory, 'Asset Allocation', year, 'Liquid Assets', month, liq_assets)
updateSpreadsheet(directory, 'Asset Allocation', year, 'Vanguard401k', month, vanguard401kbal)
driver.execute_script("window.open('https://docs.google.com/spreadsheets/d/1sWJuxtYI-fJ6bUHBWHZTQwcggd30RcOSTMlqIzd1BBo/edit#gid=953264104');")
showMessage("Balances + Review", f'MyConstant: {constant_balance} \n' f'Worthy: {worthy_balance} \n' f'Liquid Assets: {liq_assets} \n' f'401k: {vanguard401kbal}')
driver.quit()