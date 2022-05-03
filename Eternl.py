from selenium.webdriver.common.by import By
from Functions import updateSpreadsheet, getCryptocurrencyPrice

def runEternl(directory, driver):
    driver.execute_script("window.open('https://eternl.io/app/mainnet/wallet/xpub1wxalshqc32m-ml/summary');")
    # switch to last window
    driver.switch_to.window(driver.window_handles[len(driver.window_handles)-1])
    driver.implicitly_wait(10)
    adaBalance = float(driver.find_element(By.XPATH, "//*[@id='cc-main-container']/div[2]/div[1]/div/main/div[2]/div[2]/div/div/div[1]/div/div[2]/div[1]/div[2]/div[2]/div[1]").text.strip('₳').replace(',', ''))
    updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'ADA', 1, adaBalance, "ADA")
    adaPrice = getCryptocurrencyPrice('cardano')['cardano']['usd']
    updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'ADA', 2, adaPrice, "ADA")
    return adaBalance
