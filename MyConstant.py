from selenium.common.exceptions import NoSuchElementException
import time
from decimal import Decimal
import pyautogui
from Functions import showMessage, getUsername, getPassword, getOTP

def runMyConstant(directory, driver):
    driver.get("https://www.myconstant.com/log-in")
    driver.maximize_window()
    #login
    try:
        driver.find_element_by_id("lg_username").send_keys(getUsername(directory, 'My Constant'))
        driver.find_element_by_id("lg_password").send_keys(getPassword(directory, 'My Constant'))
        driver.find_element_by_id("lg_password").click()
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('space')
        showMessage("CAPTCHA", "Verify captcha, then click OK")
        driver.maximize_window()
        driver.find_element_by_xpath("//*[@id='submit-btn']").click()
        token = getOTP('my_constant')
        char = 0
        time.sleep(2)
        while char < 6:
            xpath_start = "//*[@id='layout']/div[2]/div/div/div[2]/div/div/div/div[3]/div/div/div["
            driver.find_element_by_xpath(xpath_start + str(char + 1) + "]/input").send_keys(token[char])
            char += 1
        time.sleep(6)
    except NoSuchElementException:
        exception = "caught"
    pyautogui.moveTo(1650, 167)
    pyautogui.moveTo(1670, 167)
    pyautogui.moveTo(1650, 167)
    time.sleep(8)
    # capture and format Bonds balance
    constant_balance_raw = driver.find_element_by_id("acc_balance").text.strip('$').replace(',','')
    constant_balance_dec = Decimal(constant_balance_raw)
    constant_balance = float(round(constant_balance_dec, 2))
    driver.get('https://www.myconstant.com/lend-crypto-to-earn-interest')
    time.sleep(1)
    # get BTC balance
    # click dropdown menu
    driver.find_element_by_xpath("//*[@id='layout']/div[2]/div/div/div/div[2]/div/form/div[1]/div[2]/div/div/button/div").click()
    # search for coin
    driver.find_element_by_id('dropdown-search-selectedSymbol').send_keys('BTC')
    # select coin
    driver.find_element_by_xpath("//*[@id='layout']/div[2]/div/div/div/div[2]/div/form/div[1]/div[2]/div/div/div/a/div/div/span").click()
    time.sleep(6)
    btc_balance = driver.find_element_by_xpath('/html/body/div[1]/div[1]/div[2]/div/div/div/div[2]/div/form/div[2]/div[2]/span/span/span').text
    # get ETH balance
     # click dropdown menu
    driver.find_element_by_xpath("//*[@id='layout']/div[2]/div/div/div/div[2]/div/form/div[1]/div[2]/div/div/button/div").click()
    # search for coin
    driver.find_element_by_id('dropdown-search-selectedSymbol').send_keys('ETH')
    # select coin
    driver.find_element_by_xpath("//*[@id='layout']/div[2]/div/div/div/div[2]/div/form/div[1]/div[2]/div/div/div/a/div/div/span").click()
    time.sleep(6)
    eth_balance = driver.find_element_by_xpath('/html/body/div[1]/div[1]/div[2]/div/div/div/div[2]/div/form/div[2]/div[2]/span/span/span').text
    return [constant_balance, float(btc_balance), float(eth_balance)]