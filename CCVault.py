def runCCVault(driver):
    driver.execute_script("window.open('https://ccvault.io/app/mainnet/wallet/xpub1wxalshqc32m-ml/summary');")
    # switch to last window
    driver.switch_to.window(driver.window_handles[len(driver.window_handles)-1])
    driver.implicitly_wait(6)
    ada_balance = driver.find_element_by_xpath("//*[@id='q-app']/div/div[3]/div[2]/div[1]/div/main/div[2]/div[2]/div/div/div[1]/div/div/div[4]/div[1]/div[2]/div[1]").text.strip('₳').strip()
    return float(ada_balance)