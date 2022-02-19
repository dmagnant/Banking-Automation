import time
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time

def runPresearch(driver):    
    driver.execute_script("window.open('https://nodes.presearch.org/dashboard');")
    # switch to last window
    driver.switch_to.window(driver.window_handles[len(driver.window_handles)-1])
    avail_to_stake = float(driver.find_element(By.XPATH, '/html/body/div[2]/div[2]/div[3]/div[1]/div[2]/div/div[2]/div/h2').text.strip(' PRE'))
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
        avail_to_stake = float(driver.find_element(By.XPATH, '/html/body/div[2]/div[2]/div[3]/div[1]/div[2]/div/div[2]/div/h2').text.strip(' PRE'))
    
    # stake available PRE to highest rated node
    time.sleep(2)
    if avail_to_stake:
        # get reliability scores
        num = 1
        node_found = False
        while not node_found:
            name = driver.find_element(By.XPATH, '/html/body/div[2]/div[2]/div[3]/div[5]/div/table/tbody/tr[' + str(num) + ']/td[1]/a[1]').text
            if name.lower() == 'digital ocean':
                stake_amount = avail_to_stake
                driver.find_element(By.XPATH, '/html/body/div[2]/div[2]/div[3]/div[5]/div/table/tbody/tr[' + str(num) + ']/td[9]/a[1]').click()
                while stake_amount > 0:
                    driver.find_element(By.ID, 'stake_amount').send_keys(Keys.ARROW_UP)
                    stake_amount -= 1
                driver.find_element(By.XPATH, "//*[@id='editNodeForm']/div[7]/button").click()
                time.sleep(2)
                driver.get('https://nodes.presearch.org/dashboard')
                node_found = True
            num += 1
    search_rewards = float(driver.find_element(By.XPATH, '/html/body/div[1]/header/div[2]/div[2]/div/div[1]/div/div[1]/div/span[1]').text.strip(' PRE'))
    staked_tokens = float(driver.find_element(By.XPATH, '/html/body/div[2]/div[2]/div[3]/div[1]/div[2]/div/div[1]/div/h2').text.strip(' PRE').replace(',', ''))
    pre_total = search_rewards + staked_tokens
    return [pre_total, staked_tokens]
