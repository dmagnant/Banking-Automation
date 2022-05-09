import time
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from Functions import updateSpreadsheet, showMessage, getCryptocurrencyPrice, setDirectory, chromeDriverAsUser

def login(directory, driver):
    driver.get("https://midas.investments/?login=true")
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
    
    updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'BTC_midas', 1, balances[0], "BTC")
    btcPrice = getCryptocurrencyPrice('bitcoin')['bitcoin']['usd']
    updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'BTC_midas', 2, btcPrice, "BTC")

    updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'ETH_midas', 1, balances[1], "ETH")
    ethPrice = getCryptocurrencyPrice('ethereum')['ethereum']['usd']
    updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'ETH_midas', 2, ethPrice, "ETH")

    return balances

if __name__ == '__main__':
    directory = setDirectory()
    driver = chromeDriverAsUser()
    driver.implicitly_wait(3)
    response = runMidas(directory, driver)
    print('btc balance: ' + response[0])
    print('eth balance: ' + response[1])

