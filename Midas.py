import time

from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By

from Functions import (chromeDriverAsUser, getCryptocurrencyPrice,
                       setDirectory, showMessage,
                       updateCoinQuantityFromStakingInGnuCash,
                       updateCryptoPriceInGnucash, updateSpreadsheet)


def login(directory, driver):
    driver.execute_script("window.open('https://midas.investments/?login=true');")
    # switch to last window
    driver.switch_to.window(driver.window_handles[len(driver.window_handles)-1])
    #login
    try: 
        driver.find_element(By.XPATH, "//*[@id='app-dialogs-root']/div[1]/div/div/button[1]")
        showMessage('email 2fa', 'get code from email, paste in window and confirm, then click OK')
        time.sleep(3)
    except NoSuchElementException:
        exception = "already logged in" 


def captureBalances(driver):
    driver.get("https://midas.investments/assets")
    btcBalance = float(driver.find_element(By.XPATH, "//*[@id='ga-asset-table-invest-BTC']/span[2]/span").text.replace('BTC', ''))
    ethBalance = float(driver.find_element(By.XPATH, "//*[@id='ga-asset-table-invest-ETH']/span[2]/span").text.replace('ETH', ''))
    return [btcBalance, ethBalance]


def runMidas(directory, driver):
    login(directory, driver)
    balances = captureBalances(driver)
    
    updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'BTC-Midas', 1, balances[0], "BTC")
    updateCoinQuantityFromStakingInGnuCash(balances[0], 'BTC-Midas')
    btcPrice = getCryptocurrencyPrice('bitcoin')['bitcoin']['usd']
    updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'BTC-Midas', 2, btcPrice, "BTC")
    updateCryptoPriceInGnucash('BTC', format(btcPrice, ".2f"))

    updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'ETH-Midas', 1, balances[1], "ETH")
    updateCoinQuantityFromStakingInGnuCash(balances[1], 'ETH-Midas')
    ethPrice = getCryptocurrencyPrice('ethereum')['ethereum']['usd']
    updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'ETH-Midas', 2, ethPrice, "ETH")
    updateCryptoPriceInGnucash('ETH', format(ethPrice, ".2f"))

    return balances

if __name__ == '__main__':
    directory = setDirectory()
    driver = chromeDriverAsUser()
    driver.implicitly_wait(3)
    response = runMidas(directory, driver)
    print('btc balance: ' + str(response[0]))
    print('eth balance: ' + str(response[1]))

