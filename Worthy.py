from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
import time
from decimal import Decimal
from Functions import getUsername, getPassword, setDirectory, chromeDriverAsUser

def login(directory, driver):
    driver.execute_script("window.open('https://worthy.capital/start');")
    # switch to last window
    driver.switch_to.window(driver.window_handles[len(driver.window_handles)-1])
    time.sleep(3)
    try:
        # login
        # click Login button
        driver.find_element(By.XPATH, "//*[@id='q-app']/div/div[1]/div/div[2]/div/button[2]/span[2]/span").click()
        # enter credentials
        driver.find_element(By.XPATH, "//*[@id='auth0-lock-container-1']/div/div[2]/form/div/div/div[2]/span/div/div/div/div/div/div/div/div/div[2]/div[1]/div[1]/input").send_keys(getUsername(directory, 'Worthy'))
        driver.find_element(By.XPATH, "//*[@id='auth0-lock-container-1']/div/div[2]/form/div/div/div[2]/span/div/div/div/div/div/div/div/div/div[2]/div[2]/div/div[1]/input").send_keys(getPassword(directory, 'Worthy'))
        # click Login button (again)
        driver.find_element(By.XPATH, "//*[@id='auth0-lock-container-1']/div/div[2]/form/div/div/button/span").click()
    except NoSuchElementException:
        exception = "caught"
    time.sleep(3)


def captureBalance(driver):
    # Get balance from Worthy I
    driver.find_element(By.XPATH, "//*[@id='q-app']/div/div[1]/main/div/div/div[2]/div/div[2]/div/div/div[1]").click()
    wpc1Raw = driver.find_element(By.XPATH, "/html/body/div[1]/div/div[1]/main/div/div/div[2]/div/div[2]/div/div/div[3]/div/h4/span[3]").text.strip('$').replace(',','')
    wpc1 = Decimal(wpc1Raw)
    # Get balance from Worthy II
    driver.find_element(By.XPATH, "/html/body/div[1]/div/div[1]/main/div/div/div[2]/div/div[3]/div/div/div[2]/div/h4/span[3]").click()
    wpc2Raw = driver.find_element(By.XPATH, "/html/body/div[1]/div/div[1]/main/div/div/div[2]/div/div[3]/div/div/div[2]/div/h4/span[3]").text.strip('$').replace(',','')
    wpc2 = Decimal(wpc2Raw)
    # Combine Worthy balances
    worthyBalanceDec = wpc1 + wpc2
    # convert from decimal to float
    worthyBalance = float(worthyBalanceDec)
    return worthyBalance

def runWorthy(directory, driver):    
    login(directory, driver)
    return captureBalance(driver)

if __name__ == '__main__':
    directory = setDirectory()
    driver = chromeDriverAsUser()
    response = runWorthy(directory, driver)
    print('balance: ' + response)