import time
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from Functions import getUsername, getPassword, getOTP, updateSpreadsheet, getCryptocurrencyPrice, setDirectory, chromeDriverAsUser

def runKraken(directory, driver):    
    driver.execute_script("window.open('https://www.kraken.com/sign-in');")
    driver.implicitly_wait(2)
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
    driver.get('https://www.kraken.com/u/history/ledger')
    eth2Balance = ''
    solBalance = ''
    dotBalance = ''
    algoBalance = ''
    num = 1
    while num < 20:
        balance = driver.find_element(By.XPATH, "//*[@id='__next']/div/main/div/div[2]/div/div/div[3]/div[2]/div/div[" + str(num) + "]/div/div[7]/div/div/span/span/span").text
        coin = driver.find_element(By.XPATH, "//*[@id='__next']/div/main/div/div[2]/div/div/div[3]/div[2]/div/div[" + str(num) + "]/div/div[7]/div/div/div").text
        if coin == 'ETH2':
            if not eth2Balance:
                eth2Balance = float(balance)
        elif coin == 'SOL':
            if not solBalance:
                solBalance = float(balance)
        elif coin == 'DOT':
            if not dotBalance:
                dotBalance = float(balance)
        elif coin == 'ALGO':
            if not algoBalance:
                algoBalance = float(balance)
        num = 21 if eth2Balance and solBalance and dotBalance and algoBalance else num + 1
    
    updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'ALGO', 1, algoBalance, "ALGO")
    algoPrice = getCryptocurrencyPrice('algorand')['algorand']['usd']
    updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'ALGO', 2, algoPrice, "ALGO")

    updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'DOT', 1, dotBalance, "DOT")
    dotPrice = getCryptocurrencyPrice('polkadot')['polkadot']['usd']
    updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'DOT', 2, dotPrice, "DOT")

    updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'ETH2', 1, eth2Balance, "ETH2")
    eth2Price = getCryptocurrencyPrice('ethereum')['ethereum']['usd']
    updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'ETH_kraken', 2, eth2Price, "ETH")
    updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'ETH2', 2, eth2Price, "ETH2")


    updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'SOL', 1, solBalance, "SOL")
    solPrice = getCryptocurrencyPrice('solana')['solana']['usd']
    updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'SOL', 2, solPrice, "SOL")

    return [algoBalance, dotBalance, eth2Balance, solBalance]

if __name__ == '__main__':
    directory = setDirectory()
    driver = chromeDriverAsUser()
    response = runKraken(directory, driver)
    print('algo balance: ' + response[0])
    print('dot balance: ' + response[1])
    print('eth2 balance: ' + response[2])
    print('sol balance: ' + response[3])
