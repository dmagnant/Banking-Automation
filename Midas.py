import time
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from Functions import updateSpreadsheet, showMessage, getCryptocurrencyPrice

def runMidas(directory, driver):
    driver.get("https://midas.investments/?login=true")
    driver.implicitly_wait(3)
    #login
    try: 
        driver.find_element(By.XPATH, "//*[@id='app-dialogs-root']/div[1]/div/div/button[1]")
        showMessage('email 2fa', 'get code from email, paste in window and confirm, then click OK')
        time.sleep(3)
    except NoSuchElementException:
        exception = "already logged in"
    driver.get("https://midas.investments/assets")
    btcBalance = float(driver.find_element(By.XPATH, "//*[@id='ga-asset-table-invest-BTC']/span[2]/span").text.replace('BTC', ''))
    ethBalance = float(driver.find_element(By.XPATH, "//*[@id='ga-asset-table-invest-ETH']/span[2]/span").text.replace('ETH', ''))
    updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'BTC_midas', 1, btcBalance, "BTC")
    btcPrice = getCryptocurrencyPrice('bitcoin')['bitcoin']['usd']
    updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'BTC_midas', 2, btcPrice, "BTC")

    updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'ETH_midas', 1, ethBalance, "ETH")
    ethPrice = getCryptocurrencyPrice('ethereum')['ethereum']['usd']
    updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'ETH_midas', 2, ethPrice, "ETH")

    return [btcBalance, ethBalance]
