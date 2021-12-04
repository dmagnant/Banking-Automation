from datetime import datetime
from decimal import Decimal
from Worthy import runWorthy
from MyConstant import runMyConstant
from HealthEquity import runHealthEquity
from Functions import setDirectory, chromeDriverAsUser, openGnuCashBook, showMessage, getGnuCashBalance, updateSpreadsheet, getStartAndEndOfPreviousMonth, writeGnuTransaction

directory = setDirectory()
driver = chromeDriverAsUser(directory)
driver.implicitly_wait(6)
my_constant_balances = runMyConstant(directory, driver)
constant_balance = my_constant_balances[0]
worthy_balance = runWorthy(directory, driver)
# calculate total bond balance
bond_balance = worthy_balance + constant_balance
# Set Gnucash Book
mybook = openGnuCashBook(directory, 'Finance', False, False)
# Obtain balances from Gnucash
constant_old_bal = float(getGnuCashBalance(mybook, 'MyConstant'))
worthy_old_bal = float(getGnuCashBalance(mybook, 'Worthy'))
# Calculate Interest
constant_interest = Decimal(constant_balance - constant_old_bal)
worthy_interest = Decimal(worthy_balance - worthy_old_bal)
# Get current date
today = datetime.today()
year = today.year
month = today.month
lastmonth = getStartAndEndOfPreviousMonth(today, month, year)
# Add Interest Transaction to My Constant
writeGnuTransaction(mybook, "Interest", lastmonth[1], -round(constant_interest, 2), "Income:Investments:Interest", "Assets:Liquid Assets:My Constant")
# Add Interest Transaction to Worthy Bonds
writeGnuTransaction(mybook, "Interest", lastmonth[1], -round(worthy_interest, 2), "Income:Investments:Interest", "Assets:Liquid Assets:Worthy Bonds")
liq_assets = getGnuCashBalance(mybook, 'Liquid Assets')
HE_balances = runHealthEquity(driver, lastmonth)
HE_hsa_balance = HE_balances[0]
HE_hsa_dividends = HE_balances[1]
vanguard401kbal = HE_balances[2]
# capture GnuCash HSA Balance
HSA_gnu_balance = getGnuCashBalance(mybook, 'HSA')
# Capture HSA change
HE_hsa_change = Decimal(HE_hsa_balance - float(HSA_gnu_balance))
# Capture market change
HE_hsa_mkt_change = Decimal(HE_hsa_change - HE_hsa_dividends)
# format data
HE_hsa_change = round(HE_hsa_change, 2)
HE_hsa_mkt_change = round(HE_hsa_mkt_change, 2)
writeGnuTransaction(mybook, "HSA Statement", lastmonth[1], [HE_hsa_change, -HE_hsa_dividends, -HE_hsa_mkt_change], ["Income:Investments:Dividends", "Income:Investments:Market Change"], "Assets:Non-Liquid Assets:HSA")
updateSpreadsheet(directory, 'Asset Allocation', year, 'Bonds', month, bond_balance)
updateSpreadsheet(directory, 'Asset Allocation', year, 'HE_HSA', month, HE_hsa_balance)
updateSpreadsheet(directory, 'Asset Allocation', year, 'Liquid Assets', month, float(liq_assets))
updateSpreadsheet(directory, 'Asset Allocation', year, 'Vanguard401k', month, vanguard401kbal)
updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'BTC', 1, my_constant_balances[1])
updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'ETH', 1, my_constant_balances[2])

driver.execute_script("window.open('https://docs.google.com/spreadsheets/d/1sWJuxtYI-fJ6bUHBWHZTQwcggd30RcOSTMlqIzd1BBo/edit#gid=953264104');")
showMessage("Balances + Review", f'MyConstant: {constant_balance} \n' f'Worthy: {worthy_balance} \n' f'Liquid Assets: {liq_assets} \n' f'401k: {vanguard401kbal}')
driver.quit()