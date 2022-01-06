import time
from selenium.common.exceptions import NoSuchElementException
import time
from decimal import Decimal
from Functions import getUsername, getPassword

def runWorthy(directory, driver):    
    driver.execute_script("window.open('https://worthy.capital/start');")
    # switch to last window
    driver.switch_to.window(driver.window_handles[len(driver.window_handles)-1])
    time.sleep(2)
    try:
        # login
        # click Login button
        driver.find_element_by_xpath("//*[@id='q-app']/div/div[1]/div/div[2]/div/button[2]/span[2]/span").click()
        # enter credentials
        driver.find_element_by_xpath("//*[@id='auth0-lock-container-1']/div/div[2]/form/div/div/div[2]/span/div/div/div/div/div/div/div/div/div[2]/div[1]/div[1]/input").send_keys(getUsername(directory, 'Worthy'))
        driver.find_element_by_xpath("//*[@id='auth0-lock-container-1']/div/div[2]/form/div/div/div[2]/span/div/div/div/div/div/div/div/div/div[2]/div[2]/div/div[1]/input").send_keys(getPassword(directory, 'Worthy'))
        # click Login button (again)
        driver.find_element_by_xpath("//*[@id='auth0-lock-container-1']/div/div[2]/form/div/div/button/span").click()
    except NoSuchElementException:
        exception = "caught"
    # capture balances
    time.sleep(3)
    # Get balance from Worthy I

    driver.find_element_by_xpath("//*[@id='q-app']/div/div[1]/main/div/div/div[2]/div/div[2]/div/div/div[1]").click()
    wpc1_raw = driver.find_element_by_xpath("/html/body/div[1]/div/div[1]/main/div/div/div[2]/div/div[2]/div/div/div[3]/div/h4/span[3]").text.strip('$').replace(',','')
    wpc1 = Decimal(wpc1_raw)
    # Get balance from Worthy II
    driver.find_element_by_xpath("/html/body/div[1]/div/div[1]/main/div/div/div[2]/div/div[3]/div/div/div[2]/div/h4/span[3]").click()
    wpc2_raw = driver.find_element_by_xpath("/html/body/div[1]/div/div[1]/main/div/div/div[2]/div/div[3]/div/div/div[2]/div/h4/span[3]").text.strip('$').replace(',','')
    wpc2 = Decimal(wpc2_raw)
    # Combine Worthy balances
    worthy_balance_dec = wpc1 + wpc2
    # convert from decimal to float
    worthy_balance = float(worthy_balance_dec)
    return worthy_balance