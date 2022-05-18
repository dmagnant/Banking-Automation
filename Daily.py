import os
import os.path
from Ally import runAlly
from M1Finance import runM1
from TIAA import runTIAA
from Presearch import runPresearch

from Functions import setDirectory, chromeDriverAsUser, openGnuCashBook, getGnuCashBalance, showMessage, startExpressVPN, updateCryptoPrices

directory = setDirectory()
driver = chromeDriverAsUser(directory)
driver.implicitly_wait(3)
m1 = runM1(directory, driver)
tiaa = runTIAA(directory, driver)
stakedPre = runPresearch(directory, driver)
preToGo = abs((stakedPre[1] % 2000) - 2000)
Finance = openGnuCashBook(directory, 'Finance', True, True)
m1Gnu = getGnuCashBalance(Finance, 'M1')
tiaaGnu = getGnuCashBalance(Finance, 'TIAA')
updateCryptoPrices()
cryptoBalance = getGnuCashBalance(Finance, 'Crypto')
driver.execute_script("window.open('https://docs.google.com/spreadsheets/d/1684fQ-gW5A0uOf7s45p9tC4GiEE5s5_fjO5E7dgVI1s/edit#gid=1688093622');")
if m1[1]:
    os.startfile(directory + r"\Finances\Personal Finances\Finance.gnucash")
showMessage("Balances + Review", 
    f'M1 Spend Balance: {m1[0]} \n' 
    f'GnuCash Balance: {m1Gnu} \n \n'
    f'TIAA: {tiaa} \n'
    f'GnuCash Balance: {tiaaGnu} \n \n'
    f'Crypto Balance: {cryptoBalance} \n \n'
    f'{preToGo} PRE until next Node \n \n'
    f'Review transactions:\n {m1[1]}')
driver.quit()
driver = chromeDriverAsUser(directory)
ally = runAlly(directory, driver)
Home = openGnuCashBook(directory, 'Home', True, True)
allyGnu = getGnuCashBalance(Home, 'Ally')
driver.execute_script("window.open('https://docs.google.com/spreadsheets/d/1oP3U7y8qywvXG9U_zYXgjFfqHrCyPtUDl4zPDftFCdM/edit#gid=317262693');")
if ally[1]:
    os.startfile(directory + r"\Stuff\Home\Finances\Home.gnucash")
showMessage("Balances + Review", f'Ally Balance: {ally[0]} \n'f'GnuCash Balance: {allyGnu} \n \n'f'Review transactions:\n {ally[1]}')
driver.quit()
startExpressVPN()