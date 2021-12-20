import time
from Functions import closeExpressVPN
from Functions import getPassword, showMessage, closeExpressVPN

def runTIAA(directory, driver):
    closeExpressVPN()
    # # # TIAA
    driver.execute_script("window.open('https://www.tiaabank.com/');")
    # switch to last window
    driver.switch_to.window(driver.window_handles[len(driver.window_handles)-1])
    ## Login
    driver.find_element_by_id("password").send_keys(getPassword(directory, 'TIAA'))
    # click LOG IN
    driver.find_element_by_xpath("/html/body/tb-app/tb-tmpl/div/section[1]/div/article/div[2]/tb-login-personal/div/div[1]/form/div/div[3]/button").click()
    # handle captcha
    showMessage('CAPTCHA', "Verify captcha, then click OK")
    time.sleep(3)
    driver.get("https://shared.tiaa.org/private/banktxns/tiaabank?number=f2745d23d777a9b7f1378c1cb00d36c2")
    tiaa = driver.find_element_by_xpath("//*[@id='hero-balance']/div[1]/div[2]").text.strip('$').strip(',')
    return tiaa