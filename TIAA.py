import time

from selenium.webdriver.common.by import By

from Functions import (chromeDriverAsUser, closeExpressVPN, getPassword,
                       setDirectory, showMessage)


def login(directory, driver):
    closeExpressVPN()
    time.sleep(3)
    # # # TIAA
    driver.execute_script("window.open('https://www.tiaabank.com/');")
    # switch to last window
    driver.switch_to.window(driver.window_handles[len(driver.window_handles)-1])
    ## Login 
    # driver.find_element(By.ID, "password").send_keys(getPassword(directory, 'TIAA'))   # password pre-filled
    time.sleep(1)
    # click LOG IN
    driver.find_element(By.XPATH, "/html/body/tb-app/tb-tmpl/div/section[1]/div/article/div[2]/tb-login-personal/div/div[1]/form/div/div[3]/button").click()
    # handle captcha
    showMessage('CAPTCHA', "Verify captcha, then click OK")
    time.sleep(3)

def captureBalance(driver):
    driver.get("https://shared.tiaa.org/private/banktxns/tiaabank?number=f2745d23d777a9b7f1378c1cb00d36c2")
    return driver.find_element(By.XPATH, "//*[@id='hero-balance']/div[1]/div[2]").text.strip('$').replace(',','')

def runTIAA(directory, driver):
    login(directory, driver)
    return captureBalance(driver)

if __name__ == '__main__':
    directory = setDirectory()
    driver = chromeDriverAsUser()
    response = runTIAA(directory, driver)
    print('balance: ' + response)
    