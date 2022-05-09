import time
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
from Functions import updateSpreadsheet, setDirectory, getCryptocurrencyPrice, chromeDriverAsUser


def claimRewards(driver):
    driver.execute_script("window.open('https://nodes.presearch.org/dashboard');")
    # switch to last window
    driver.switch_to.window(driver.window_handles[len(driver.window_handles)-1])
    availToStake = float(driver.find_element(By.XPATH, '/html/body/div[2]/div[2]/div[3]/div[1]/div[2]/div/div[2]/div/h2').text.strip(' PRE'))
    # claim rewards
    unclaimed = driver.find_element(By.XPATH, '/html/body/div[2]/div[2]/div[3]/div[2]/div[3]/div[2]/div/div/div[1]/h2').text.strip(' PRE')
    if float(unclaimed) > 0:
        # click Claim
        driver.find_element(By.XPATH, '/html/body/div[2]/div[2]/div[3]/div[2]/div[3]/div[2]/div/div/div[2]/div/a').click()
        # click Claim Reward
        driver.find_element(By.XPATH, '/html/body/div[2]/div[2]/div/div/div/div[2]/div/form/div/button').click()
        time.sleep(4)
        driver.refresh()
        time.sleep(1)
        availToStake = float(driver.find_element(By.XPATH, '/html/body/div[2]/div[2]/div[3]/div[1]/div[2]/div/div[2]/div/h2').text.strip(' PRE'))
    
    # stake available PRE to highest rated node
    time.sleep(2)
    if availToStake:
        # get reliability scores
        num = 1
        node_found = False
        while not node_found:
            name = driver.find_element(By.XPATH, '/html/body/div[2]/div[2]/div[3]/div[5]/div/table/tbody/tr[' + str(num) + ']/td[1]/a[1]').text
            if name.lower() == 'aws':
                stakeAmount = availToStake
                driver.find_element(By.XPATH, '/html/body/div[2]/div[2]/div[3]/div[5]/div/table/tbody/tr[' + str(num) + ']/td[11]/a[1]').click()
                while stakeAmount > 0:
                    driver.find_element(By.ID, 'stake_amount').send_keys(Keys.ARROW_UP)
                    stakeAmount -= 1
                driver.find_element(By.XPATH, "//*[@id='editNodeForm']/div[7]/button").click()
                time.sleep(2)
                driver.get('https://nodes.presearch.org/dashboard')
                node_found = True
            num += 1


def captureBalance(driver):
    searchRewards = float(driver.find_element(By.XPATH, '/html/body/div[1]/header/div[2]/div[2]/div/div[1]/div/div[1]/div/span[1]').text.strip(' PRE'))
    stakedTokens = float(driver.find_element(By.XPATH, '/html/body/div[2]/div[2]/div[3]/div[1]/div[2]/div/div[1]/div/h2').text.strip(' PRE').replace(',', ''))
    preTotal = searchRewards + stakedTokens
    return [preTotal, stakedTokens]
     

def runPresearch(directory, driver):    
    claimRewards(driver)
    balances = captureBalance(driver)

    updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'PRE', 1, balances[0], "PRE")
    prePrice = getCryptocurrencyPrice('presearch')['presearch']['usd']
    updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'PRE', 2, prePrice, "PRE")

    return balances
    
    
if __name__ == '__main__':
    directory = setDirectory()
    driver = chromeDriverAsUser()
    response = runPresearch(directory, driver)
    print('balance: ' + response[0])
