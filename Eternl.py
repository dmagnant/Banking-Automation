import time
from selenium.webdriver.common.by import By
from Functions import updateSpreadsheet, getCryptocurrencyPrice, setDirectory, chromeDriverAsUser, updateCryptoPrice

def runEternl(directory, driver):
    driver.execute_script("window.open('https://eternl.io/app/mainnet/wallet/xpub1wxalshqc32m-ml/summary');")
    # switch to last window
    driver.switch_to.window(driver.window_handles[len(driver.window_handles)-1])
    driver.implicitly_wait(10)
    adaBalance = float(driver.find_element(By.XPATH, "//*[@id='cc-main-container']/div/div[3]/div[2]/nav/div/div[2]/div/div/div/div[3]").text.strip('(initializing)').replace('\n', '').strip('₳').replace(',', ''))
    updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'ADA', 1, adaBalance, "ADA")
    adaPrice = getCryptocurrencyPrice('cardano')['cardano']['usd']
    updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'ADA', 2, adaPrice, "ADA")
    updateCryptoPrice('ADA', format(adaPrice, ".2f"))
    return adaBalance

if __name__ == '__main__':
    directory = setDirectory()
    driver = chromeDriverAsUser()
    response = runEternl(directory, driver)
    print('balance: ' + str(response))
