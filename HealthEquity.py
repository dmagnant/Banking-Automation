from selenium.common.exceptions import NoSuchElementException
import time
from selenium.webdriver.common.keys import Keys
from decimal import Decimal
from datetime import datetime
from Functions import showMessage


def runHealthEquity(driver, lastmonth):
    # # # Health Equity HSA
    driver.execute_script("window.open('https://member.my.healthequity.com/hsa/21895515-010');")
    # switch to last window
    driver.switch_to.window(driver.window_handles[len(driver.window_handles)-1])
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
    return [HE_hsa_balance, HE_hsa_dividends, vanguard401kbal]
