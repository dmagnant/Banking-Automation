from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
import time
from decimal import Decimal
import pyautogui
from Functions import showMessage, getUsername, getPassword, getOTP, updateSpreadsheet, getCryptocurrencyPrice, setDirectory, chromeDriverAsUser, updateCryptoPrice


def login(directory, driver):
    driver.execute_script("window.open('https://www.myconstant.com/log-in');")
    # switch to last window
    driver.switch_to.window(driver.window_handles[len(driver.window_handles)-1])
    #login
    try:
        driver.find_element(By.ID, "lg_username").send_keys(getUsername(directory, 'My Constant'))
        driver.find_element(By.ID, "lg_password").send_keys(getPassword(directory, 'My Constant'))
        driver.find_element(By.ID, "lg_password").click()
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('space')
        showMessage("CAPTCHA", "Verify captcha, then click OK")
        driver.find_element(By.XPATH, "//*[@id='submit-btn']").click()
        token = getOTP('my_constant')
        char = 0
        time.sleep(2)
        while char < 6:
            xpath_start = "//*[@id='layout']/div[2]/div/div/div[2]/div/div/div/div[3]/div/div/div["
            driver.find_element(By.XPATH, xpath_start + str(char + 1) + "]/input").send_keys(token[char])
            char += 1
        time.sleep(6)
    except NoSuchElementException:
        exception = "caught"


def getCoinBalance(driver, coin):
        # click dropdown menu
        driver.find_element(By.XPATH, "//*[@id='layout']/div[2]/div/div/div/div[2]/div/form/div[1]/div[2]/div/div/button/div").click()
        # search for coin
        driver.find_element(By.ID, 'dropdown-search-selectedSymbol').send_keys(coin)
        # select coin
        driver.find_element(By.XPATH, "//*[@id='layout']/div[2]/div/div/div/div[2]/div/form/div[1]/div[2]/div/div/div/a/div/div[1]").click()
        time.sleep(6)
        return driver.find_element(By.XPATH, "//*[@id='layout']/div[2]/div/div/div/div[2]/div/form/div[2]/div[2]/span/span/span").text


def captureBalances(driver):
    pyautogui.moveTo(1650, 167)
    pyautogui.moveTo(1670, 167)
    pyautogui.moveTo(1650, 167)
    time.sleep(8)
    # capture and format Bonds balance
    constantBalanceRaw = driver.find_element(By.ID, "acc_balance").text.strip('$').replace(',','')
    constantBalanceDec = Decimal(constantBalanceRaw)
    constantBalance = float(round(constantBalanceDec, 2))
    driver.get('https://www.myconstant.com/lend-crypto-to-earn-interest')
    time.sleep(2)
    # get coin balances
    btcBalance = float(getCoinBalance(driver, ('BTC')))
    ethBalance = float(getCoinBalance(driver, ('ETHEREUM')))

    return [constantBalance, btcBalance, ethBalance]


def runMyConstant(directory, driver):
    login(directory, driver)
    balances = captureBalances(driver)
    
    updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'BTC_myconstant', 1, balances[1], "BTC")
    btcPrice = getCryptocurrencyPrice('bitcoin')['bitcoin']['usd']
    updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'BTC_myconstant', 2, btcPrice, "BTC")
    updateCryptoPrice('BTC', format(btcPrice, ".2f"))

    updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'ETH_myconstant', 1, balances[2], "ETH")
    ethPrice = getCryptocurrencyPrice('ethereum')['ethereum']['usd']
    updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'ETH_myconstant', 2, ethPrice, "ETH")
    updateCryptoPrice('ETH', format(ethPrice, ".2f"))

    return balances

if __name__ == '__main__':
    directory = setDirectory()
    driver = chromeDriverAsUser()
    response = runMyConstant(directory, driver)
    print('myconstant balance: ' + response[0])
    print('btc balance: ' + response[1])
    print('eth balance: ' + response[2])
