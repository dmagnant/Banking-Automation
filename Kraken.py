import time
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from Functions import getUsername, getPassword, getOTP, updateSpreadsheet, getCryptocurrencyPrice, setDirectory, chromeDriverAsUser, updateCryptoPriceInGnucash, updateCoinQuantityFromStakingInGnuCash

def login(directory, driver):
    driver.execute_script("window.open('https://www.kraken.com/sign-in');")
    # switch to last window
    driver.switch_to.window(driver.window_handles[len(driver.window_handles)-1])
    time.sleep(2)
    try:
        driver.find_element(By.ID, 'username').send_keys(getUsername(directory, 'Kraken'))
        time.sleep(1)
        driver.find_element(By.ID, 'password').send_keys(getPassword(directory, 'Kraken'))
        time.sleep(1)
        driver.find_element(By.XPATH, "/html/body/div[1]/div[2]/div[2]/form/div/div[3]/button/div/div/div").click()   
        token = getOTP('kraken_otp')
        time.sleep(1)
        driver.find_element(By.ID, 'tfa').send_keys(token)
        driver.find_element(By.XPATH, "/html/body/div/div[2]/div[2]/form/div[1]/div/div/div[2]/button/div/div/div").click()
    except NoSuchElementException:
        exception = 'already logged in'
    time.sleep(2)


def captureBalances(driver):
    driver.get('https://www.kraken.com/u/history/ledger')
    algoBalance = ''
    dotBalance = ''
    eth2Balance = ''
    solBalance = ''
    num = 1
    while num < 20:
        balance = driver.find_element(By.XPATH, "//*[@id='__next']/div/main/div/div[2]/div/div/div[3]/div[2]/div/div[" + str(num) + "]/div/div[7]/div/div/span/span/span").text
        coin = driver.find_element(By.XPATH, "//*[@id='__next']/div/main/div/div[2]/div/div/div[3]/div[2]/div/div[" + str(num) + "]/div/div[7]/div/div/div").text
        if coin == 'ALGO':
            if not algoBalance:
                algoBalance = float(balance)
        elif coin == 'DOT':
            if not dotBalance:
                dotBalance = float(balance)
        if coin == 'ETH2':
            if not eth2Balance:
                eth2Balance = float(balance)
        elif coin == 'SOL':
            if not solBalance:
                solBalance = float(balance)
        num = 21 if eth2Balance and solBalance and dotBalance and algoBalance else num + 1
    
    return [algoBalance, dotBalance, eth2Balance, solBalance]

def runKraken(directory, driver):    
    login(directory, driver)
    balances = captureBalances(driver)
    
    updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'ALGO', 1, balances[0], "ALGO")
    updateCoinQuantityFromStakingInGnuCash(balances[0], 'ALGO')
    algoPrice = getCryptocurrencyPrice('algorand')['algorand']['usd']
    updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'ALGO', 2, algoPrice, "ALGO")
    updateCryptoPriceInGnucash('ALGO', format(algoPrice, ".2f"))

    updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'DOT', 1, balances[1], "DOT")
    updateCoinQuantityFromStakingInGnuCash(balances[1], 'DOT')
    dotPrice = getCryptocurrencyPrice('polkadot')['polkadot']['usd']
    updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'DOT', 2, dotPrice, "DOT")
    updateCryptoPriceInGnucash('DOT', format(dotPrice, ".2f"))


    updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'ETH2', 1, balances[2], "ETH2")
    updateCoinQuantityFromStakingInGnuCash(balances[2], 'ETH2')
    eth2Price = getCryptocurrencyPrice('ethereum')['ethereum']['usd']
    updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'ETH-Kraken', 2, eth2Price, "ETH")
    updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'ETH2', 2, eth2Price, "ETH2")
    updateCryptoPriceInGnucash('ETH', format(eth2Price, ".2f"))
    updateCryptoPriceInGnucash('ETH2', format(eth2Price, ".2f"))

    updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'SOL', 1, balances[3], "SOL")
    updateCoinQuantityFromStakingInGnuCash(balances[3], 'SOL')
    solPrice = getCryptocurrencyPrice('solana')['solana']['usd']
    updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'SOL', 2, solPrice, "SOL")
    updateCryptoPriceInGnucash('SOL', format(solPrice, ".2f"))

    return balances

if __name__ == '__main__':
    directory = setDirectory()
    driver = chromeDriverAsUser()
    driver.implicitly_wait(2)
    response = runKraken(directory, driver)
    print('algo balance: ' + str(response[0]))
    print('dot balance: ' + str(response[1]))
    print('eth2 balance: ' + str(response[2]))
    print('sol balance: ' + str(response[3]))
