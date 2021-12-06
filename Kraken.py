import time
from selenium.common.exceptions import NoSuchElementException
import time
from Functions import getUsername, getPassword, getOTP

def runKraken(directory, driver):    
    driver.execute_script("window.open('https://www.kraken.com/sign-in');")
    # switch to last window
    driver.switch_to.window(driver.window_handles[len(driver.window_handles)-1])
    time.sleep(1)
    try:
        driver.find_element_by_id('username').send_keys(getUsername(directory, 'Kraken'))
        driver.find_element_by_id('password').send_keys(getPassword(directory, 'Kraken'))
        driver.find_element_by_xpath("//*[@id='__next']/div[2]/div/div[2]/div/div/form/div/div[3]/button/div/div/div").click()
        token = getOTP('kraken_otp')
        driver.find_element_by_id('tfa').send_keys(token)
        driver.find_element_by_xpath("//*[@id='__next']/div[2]/div/div[2]/div/div/form/div[1]/div/div/div[2]/button/div/div/div/span").click()
    except NoSuchElementException:
        exception = 'already logged in'
    driver.get('https://www.kraken.com/u/history/ledger')
    eth2_balance = ''
    sol_balance = ''
    dot_balance = ''
    num = 1
    while num < 20:
        balance = driver.find_element_by_xpath("//*[@id='__next']/div/main/div/div[2]/div/div/div[3]/div[2]/div/div[" + str(num) + "]/div/div[7]/div/div/span/span/span").text
        coin = driver.find_element_by_xpath("//*[@id='__next']/div/main/div/div[2]/div/div/div[3]/div[2]/div/div[" + str(num) + "]/div/div[7]/div/div/div").text
        if coin == 'ETH2':
            if not eth2_balance:
                eth2_balance = float(balance)
        elif coin == 'SOL':
            if not sol_balance:
                sol_balance = float(balance)
        elif coin == 'DOT':
            if not dot_balance:
                dot_balance = float(balance)
        num = 21 if eth2_balance and sol_balance and dot_balance else num + 1
    return [eth2_balance, sol_balance, dot_balance]

