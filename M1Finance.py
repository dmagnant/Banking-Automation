from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from datetime import datetime
import time
import csv
from Functions import getUsername, getPassword, openGnuCashBook, showMessage, getDateRange, modifyTransactionDescription, compileGnuTransactions

def runM1(directory, driver):
    driver.get("https://dashboard.m1finance.com/login")
    driver.implicitly_wait(3)
    driver.maximize_window()
    # login
    # enter email
    driver.find_element(By.NAME, "username").send_keys(getUsername(directory, 'M1 Finance'))
    # enter password
    driver.find_element(By.NAME, "password").send_keys(getPassword(directory, 'M1 Finance'))
    # click Log In
    driver.find_element(By.XPATH, "//*[@id='root']/div/div/div/div[2]/div/div/form/div[4]/div/button").click()
    # handle captcha
    showMessage('CAPTCHA',"Verify captcha, then click OK")
    try: 
        # click Spend
        driver.find_element(By.XPATH, "//*[@id='root']/div/div/div/div[1]/div[2]/div/div[1]/nav/a[3]/div/div/span").click()
    except NoSuchElementException:
        # handle captcha
        showMessage('CAPTCHA',"Verify captcha, then click OK")
        # click Spend
        driver.find_element(By.XPATH, "//*[@id='root']/div/div/div/div[1]/div[2]/div/div[1]/nav/a[3]/div/div/span").click()
        
    m1_balance = driver.find_element(By.XPATH, "//*[@id='root']/div/div/div/div[2]/div/div[1]/div[2]/div/div[1]/div/h1").text.strip('$').replace(',', '')
    # get current date
    today = datetime.today()
    year = today.year

    date_range = getDateRange(today, 5)

    transaction = 1
    column = 1
    element = "//*[@id='root']/div/div/div/div[2]/div/div[2]/div/div[3]/a[" + str(transaction) + "]/div[" + str(column) + "]"
    inside_date_range = True
    m1_activity = directory + r"\Projects\Coding\Python\BankingAutomation\Resources\m1.csv"
    gnu_m1_activity = directory + r"\Projects\Coding\Python\BankingAutomation\Resources\gnu_m1.csv"
    open(m1_activity, 'w', newline='').truncate()
    open(gnu_m1_activity, 'w', newline='').truncate()
    # Set Gnucash Book
    mybook = openGnuCashBook(directory, 'Finance', False, False)
    clicked_next = False
    while inside_date_range:
        try:
            # capture m1_date
            raw_date = driver.find_element(By.XPATH, element).text
            if "," in raw_date:
                try:
                    # capture Date if Mon Day, Year format (standard for prior year)
                    m1_date = datetime.strptime(driver.find_element(By.XPATH, element).text, '%b %d, %Y').date()
                except ValueError:
                    # capture Date if Day of week, Mon Day format (standard for current year)
                    mod_date = datetime.strptime(raw_date, '%a, %b %d').date()
                    # adding year since its missing in M1 Finance
                    m1_date = mod_date.replace(year=year)
            else:
                # capture date if Mon Day format (only one transaction so far. REMOVE??)
                mod_date = datetime.strptime(raw_date, '%b %d').date()
                # adding year since its missing in M1 Finance
                m1_date = mod_date.replace(year=year)
            if str(m1_date) not in date_range:
                inside_date_range = False
            else:
                column += 1
                element = "//*[@id='root']/div/div/div/div[2]/div/div[2]/div/div[3]/a[" + str(transaction) + "]/div[" + str(column) + "]"
                description = driver.find_element(By.XPATH, element).text
                column += 2
                element = "//*[@id='root']/div/div/div/div[2]/div/div[2]/div/div[3]/a[" + str(transaction) + "]/div[" + str(column) + "]"
                amount = driver.find_element(By.XPATH, element).text.strip('+').replace('$', '').replace(',', '')
                # else:
                #     amount = driver.find_element(By.XPATH, element).text.strip('+').replace('$', '').replace(',', '')
                description = modifyTransactionDescription(description, amount)
                row = m1_date, description, amount
                csv.writer(open(m1_activity, 'a', newline='')).writerow(row)
                transaction += 1
                column = 1
                element = "//*[@id='root']/div/div/div/div[2]/div/div[2]/div/div[3]/a[" + str(transaction) + "]/div[" + str(column) + "]"
        except NoSuchElementException:
            if not clicked_next:
                # click Next
                driver.find_element(By.XPATH, "//*[@id='root']/div/div/div/div[2]/div/div[2]/div/div[2]/div/button[2]/span[2]").click()
                transaction = 1
                column = 1
                element = "//*[@id='root']/div/div/div/div[2]/div/div[2]/div/div[3]/a[" + str(transaction) + "]/div[" + str(column) + "]"
                time.sleep(2)
                clicked_next = True
            else:
                inside_date_range = False
    review_trans = compileGnuTransactions('M1', m1_activity, gnu_m1_activity, mybook, driver, directory, date_range, 0)
    return [m1_balance, review_trans]
    