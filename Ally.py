from selenium.common.exceptions import NoSuchElementException
from datetime import datetime
import time
import csv
from Functions import getPassword, closeExpressVPN, openGnuCashBook, getDateRange, modifyTransactionDescription, compileGnuTransactions

def runAlly(directory, driver):
    closeExpressVPN()
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
    # capture balance
    ally = driver.find_element_by_xpath("/html/body/div[1]/div[1]/main/div/div/div/div[2]/div/div[2]/div/table/tbody/tr/td[2]/div").text.replace('$', '').replace(',', '')
    # click Joint Checking link
    driver.find_element_by_partial_link_text("Joint Checking").click()
    time.sleep(3)

    # get current date
    today = datetime.today()

    date_range = getDateRange(today, 8)
    table = 2
    transaction = 1
    column = 1
    element = "//*[@id='form-elements-section']/section/section/table[" + str(table) + "]/tbody/tr[" + str(transaction) + "]/td[" + str(column) + "]"
    ally_activity = directory + r"\Projects\Coding\Python\BankingAutomation\Resources\ally.csv"
    gnu_ally_activity = directory + r"\Projects\Coding\Python\BankingAutomation\Resources\gnu_ally.csv"
    open(ally_activity, 'w', newline='').truncate()
    open(gnu_ally_activity, 'w', newline='').truncate()
    time.sleep(6)
    inside_date_range = True
    while inside_date_range:
        try:
            mod_date = datetime.strptime(driver.find_element_by_xpath(element).text, '%b %d, %Y').date()
            if str(mod_date) not in date_range:
                inside_date_range = False
            else:
                column += 1
                element = "//*[@id='form-elements-section']/section/section/table[" + str(table) + "]/tbody/tr[" + str(transaction) + "]/td[" + str(column) + "]/button"
                description = driver.find_element_by_xpath(element).text
                column += 1
                element = "//*[@id='form-elements-section']/section/section/table[" + str(table) + "]/tbody/tr[" + str(transaction) + "]/td[" + str(column) + "]"
                amount = driver.find_element_by_xpath(element).text.replace('$','').replace(',','')
                description = modifyTransactionDescription(description)
                row = str(mod_date), description, amount
                csv.writer(open(ally_activity, 'a', newline='')).writerow(row)
                transaction += 2
                column = 1
                element = "//*[@id='form-elements-section']/section/section/table[" + str(table) + "]/tbody/tr[" + str(transaction) + "]/td[" + str(column) + "]"
        except NoSuchElementException:
            table += 1
            column = 1
            transaction = 1
            element = "//*[@id='form-elements-section']/section/section/table[" + str(table) + "]/tbody/tr[" + str(transaction) + "]/td[" + str(column) + "]"
        except ValueError:
            inside_date_range = False
    # Set Gnucash Book
    mybook = openGnuCashBook(directory, 'Home', False, False)
    # Compare against existing transactions in GnuCash and import new ones
    review_trans = compileGnuTransactions('Ally', ally_activity, gnu_ally_activity, mybook, driver, directory, date_range, 0)
    return [ally, review_trans]