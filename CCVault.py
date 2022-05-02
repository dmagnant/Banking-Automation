from selenium.webdriver.common.by import By
from Functions import Functions

def runCCVault(directory, driver):
    driver.execute_script("window.open('https://ccvault.io/app/mainnet/wallet/xpub1wxalshqc32m-ml/summary');")
    # switch to last window
    driver.switch_to.window(driver.window_handles[len(driver.window_handles)-1])
    driver.implicitly_wait(10)
    ada_balance = float(driver.find_element(By.XPATH, '/html/body/div[1]/div/div[3]/div[2]/div[1]/div/main/div[2]/div[2]/div/div/div[1]/div/div/div[4]/div[2]/div[2]/div[1]').text.strip('₳').strip())
    Functions.updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'ADA', 1, ada_balance)
    return ada_balance
    