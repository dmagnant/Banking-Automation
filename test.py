import time
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException
from Ally import runAlly
from Functions import setDirectory, chromeDriverAsUser, closeExpressVPN, getPassword
directory = setDirectory()
driver = chromeDriverAsUser(directory)
driver.implicitly_wait(5)
driver.get("https://secure.ally.com/")
driver.maximize_window()
time.sleep(1)
# login
# enter password
driver.find_element_by_xpath("/html/body/div/div[1]/main/div/div/div/div/div[1]/form/div[2]/div/input").send_keys(getPassword(directory, 'Ally Bank'))
# click Log In
driver.find_element_by_xpath("/html/body/div/div[1]/main/div/div/div/div/div[1]/form/div[3]/button/span").click()
time.sleep(4)
# click Joint Checking link
driver.find_element_by_partial_link_text("Joint Checking").click()
