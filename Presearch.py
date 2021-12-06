import time
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
import time
from decimal import Decimal
from Functions import getUsername, getPassword, getOTP

def runPresearch(driver):    
    driver.execute_script("window.open('https://nodes.presearch.org/dashboard');")
    # switch to last window
    driver.switch_to.window(driver.window_handles[len(driver.window_handles)-1])
    avail_to_stake = float(driver.find_element_by_xpath('/html/body/div/div[1]/div[2]/div[3]/div[1]/div[2]/div/div[2]/div/h2').text.strip(' PRE'))
    # claim rewards
    unclaimed = driver.find_element_by_xpath('/html/body/div/div[1]/div[2]/div[3]/div[2]/div[3]/div[2]/div/div/div[1]/h2').text.strip(' PRE')
    if float(unclaimed) > 0:
        driver.find_element_by_xpath('/html/body/div/div[1]/div[2]/div[3]/div[2]/div[3]/div[2]/div/div/div[2]/div/a').click()
        driver.find_element_by_xpath('/html/body/div/div[1]/div[2]/div/div/div/div[2]/div/form/div/button').click()
        driver.refresh()
        time.sleep(1)
        avail_to_stake = float(driver.find_element_by_xpath('/html/body/div/div[1]/div[2]/div[3]/div[1]/div[2]/div/div[2]/div/h2').text.strip(' PRE'))
    
    # stake available PRE
    if avail_to_stake:
        # get reliability scores
        num = 1
        scores = []
        still_nodes = True
        while still_nodes:
            try:
                scores.append(driver.find_element_by_xpath('/html/body/div/div[1]/div[2]/div[3]/div[5]/div/table/tbody/tr[' + str(num) + ']/td[7]').text)
                num += 1
            except NoSuchElementException:
                still_nodes = False
        
        scores.sort(reverse=True)
        nodes = num - 1
        even_stake_amount = ((avail_to_stake - (avail_to_stake % nodes)) / nodes)
        remainder = (avail_to_stake % nodes)

        num = 1
        while num <= nodes:
            score = driver.find_element_by_xpath('/html/body/div/div[1]/div[2]/div[3]/div[5]/div/table/tbody/tr[' + str(num) + ']/td[7]').text
            stake_amount = even_stake_amount + remainder if score == scores[0] else even_stake_amount
            driver.find_element_by_xpath('/html/body/div/div[1]/div[2]/div[3]/div[5]/div/table/tbody/tr[' + str(num) + ']/td[9]/a[1]').click()
            while stake_amount > 0:
                driver.find_element_by_id('stake_amount').send_keys(Keys.ARROW_UP)
                stake_amount -= 1
            driver.find_element_by_xpath("//*[@id='editNodeForm']/div[7]/button").click()
            time.sleep(1)
            driver.get('https://nodes.presearch.org/dashboard')
            num += 1
    search_rewards = float(driver.find_element_by_xpath('/html/body/div/header/div/div[2]/div/div/div[1]/p').text.strip(' PRE'))
    staked_tokens = float(driver.find_element_by_xpath('/html/body/div/div[1]/div[2]/div[3]/div[1]/div[2]/div/div[1]/div/h2').text.strip(' PRE').replace(',', ''))
    pre_total = search_rewards + staked_tokens
    return pre_total

