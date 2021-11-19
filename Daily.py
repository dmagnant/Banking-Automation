import os
import os.path
from Ally import runAlly
from Functions import setDirectory, chromeDriverAsUser, openGnuCashBook, getGnuCashBalance, showMessage
from M1Finance import runM1
from TIAA import runTIAA

directory = setDirectory()
driver = chromeDriverAsUser(directory)
driver.implicitly_wait(3)
m1 = runM1(directory, driver)
tiaa = runTIAA(directory, driver)
# tiaa_balance = runTIAA(directory, driver)
Finance = openGnuCashBook(directory, 'Finance', True, True)
m1_gnu = getGnuCashBalance(Finance, 'M1')
tiaa_gnu = getGnuCashBalance(Finance, 'TIAA')
# promos_gnu = str(getGnuCashBalance(mybook, ))
driver.execute_script("window.open('https://docs.google.com/spreadsheets/d/1684fQ-gW5A0uOf7s45p9tC4GiEE5s5_fjO5E7dgVI1s/edit#gid=914927265');")
if m1[1]:
    os.startfile(directory + r"\Finances\Personal Finances\Finance.gnucash")
showMessage("Balances + Review", f'M1 Spend Balance: {m1[0]} \n' f'GnuCash Balance: {m1_gnu} \n \n'f'TIAA: {tiaa} \n'f'GnuCash Balance: {tiaa_gnu} \n \n'f'Review transactions:\n {m1[1]}')
driver.quit()
driver = chromeDriverAsUser(directory)
ally = runAlly(directory, driver)
Home = openGnuCashBook(directory, 'Home', True, True)
ally_gnu = getGnuCashBalance(Home, 'Ally')
driver.execute_script("window.open('https://docs.google.com/spreadsheets/d/1oP3U7y8qywvXG9U_zYXgjFfqHrCyPtUDl4zPDftFCdM/edit#gid=460564976');")
if ally[1]:
    os.startfile(directory + r"\Stuff\Home\Finances\Home.gnucash")
showMessage("Balances + Review", f'Ally Balance: {ally[0]} \n'f'GnuCash Balance: {ally_gnu} \n \n'f'Review transactions:\n {ally[1]}')
driver.quit()
