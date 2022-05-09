from datetime import datetime
from decimal import Decimal
from MyConstant import runMyConstant
from Worthy import runWorthy
from HealthEquity import runHealthEquity
from Functions import setDirectory, chromeDriverAsUser, openGnuCashBook, showMessage, getGnuCashBalance, updateSpreadsheet, getStartAndEndOfPreviousMonth, writeGnuTransaction

# Get current date
today = datetime.today()
year = today.year
month = today.month
lastMonth = getStartAndEndOfPreviousMonth(today, month, year)

directory = setDirectory()
driver = chromeDriverAsUser(directory)
driver.implicitly_wait(6)

myConstantBalances = runMyConstant(directory, driver, 'usd')
worthyBalance = runWorthy(directory, driver)
HEBalances = runHealthEquity(driver, lastMonth)

mybook = openGnuCashBook(directory, 'Finance', False, False)
constantInterest = Decimal(myConstantBalances[0] - float(getGnuCashBalance(mybook, 'MyConstant')))
worthyInterest = Decimal(worthyBalance - float(getGnuCashBalance(mybook, 'Worthy')))
HEHSAChange = round(Decimal(HEBalances[0] - float(getGnuCashBalance(mybook, 'HSA'))), 2)
HEHSAMarketChange = round(Decimal(HEHSAChange - HEBalances[1]), 2)

writeGnuTransaction(mybook, "Interest", lastMonth[1], -round(constantInterest, 2), "Income:Investments:Interest", "Assets:Liquid Assets:My Constant")
writeGnuTransaction(mybook, "Interest", lastMonth[1], -round(worthyInterest, 2), "Income:Investments:Interest", "Assets:Liquid Assets:Worthy Bonds")
writeGnuTransaction(mybook, "HSA Statement", lastMonth[1], [HEHSAChange, -HEBalances[1], -HEHSAMarketChange], ["Income:Investments:Dividends", "Income:Investments:Market Change"], "Assets:Non-Liquid Assets:HSA")

liquidAssets = getGnuCashBalance(mybook, 'Liquid Assets')

updateSpreadsheet(directory, 'Asset Allocation', year, 'Bonds', month, (worthyBalance + myConstantBalances[0]), 'Liquid Assets')
updateSpreadsheet(directory, 'Asset Allocation', year, 'HE_HSA', month, HEBalances[0], 'NM HSA')
updateSpreadsheet(directory, 'Asset Allocation', year, 'Liquid Assets', month, float(liquidAssets), 'Liquid Assets')
updateSpreadsheet(directory, 'Asset Allocation', year, 'Vanguard401k', month, HEBalances[2], '401k')

driver.execute_script("window.open('https://docs.google.com/spreadsheets/d/1sWJuxtYI-fJ6bUHBWHZTQwcggd30RcOSTMlqIzd1BBo/edit#gid=2058576150');")
showMessage("Balances + Review", f'MyConstant: {myConstantBalances[0]} \n' f'Worthy: {worthyBalance} \n' f'Liquid Assets: {liquidAssets} \n' f'401k: {HEBalances[2]}')
driver.quit()
