from datetime import datetime
from decimal import Decimal
from MyConstant import runMyConstant
from Worthy import runWorthy
from HealthEquity import runHealthEquity
from CCVault import runCCVault
from Kraken import runKraken
from Presearch import runPresearch
from Functions import setDirectory, chromeDriverAsUser, openGnuCashBook, showMessage, getGnuCashBalance, updateSpreadsheet, getStartAndEndOfPreviousMonth, writeGnuTransaction

# Get current date
today = datetime.today()
year = today.year
month = today.month
lastmonth = getStartAndEndOfPreviousMonth(today, month, year)

directory = setDirectory()
driver = chromeDriverAsUser(directory)
driver.implicitly_wait(6)

my_constant_balances = runMyConstant(directory, driver)
worthy_balance = runWorthy(directory, driver)
HE_balances = runHealthEquity(driver, lastmonth)
ada_balance = runCCVault(driver)
kraken_balances = runKraken(directory, driver)
pre_balance = runPresearch(driver)

# Set Gnucash Book
mybook = openGnuCashBook(directory, 'Finance', False, False)
constant_interest = Decimal(my_constant_balances[0] - float(getGnuCashBalance(mybook, 'MyConstant')))
worthy_interest = Decimal(worthy_balance - float(getGnuCashBalance(mybook, 'Worthy')))
liq_assets = getGnuCashBalance(mybook, 'Liquid Assets')
HE_hsa_change = round(Decimal(HE_balances[0] - float(getGnuCashBalance(mybook, 'HSA'))), 2)
HE_hsa_mkt_change = round(Decimal(HE_hsa_change - HE_balances[1]), 2)

writeGnuTransaction(mybook, "Interest", lastmonth[1], -round(constant_interest, 2), "Income:Investments:Interest", "Assets:Liquid Assets:My Constant")
writeGnuTransaction(mybook, "Interest", lastmonth[1], -round(worthy_interest, 2), "Income:Investments:Interest", "Assets:Liquid Assets:Worthy Bonds")
writeGnuTransaction(mybook, "HSA Statement", lastmonth[1], [HE_hsa_change, -HE_balances[1], -HE_hsa_mkt_change], ["Income:Investments:Dividends", "Income:Investments:Market Change"], "Assets:Non-Liquid Assets:HSA")

updateSpreadsheet(directory, 'Asset Allocation', year, 'Bonds', month, (worthy_balance + my_constant_balances[0]))
updateSpreadsheet(directory, 'Asset Allocation', year, 'HE_HSA', month, HE_balances[0])
updateSpreadsheet(directory, 'Asset Allocation', year, 'Liquid Assets', month, float(liq_assets))
updateSpreadsheet(directory, 'Asset Allocation', year, 'Vanguard401k', month, HE_balances[2])
updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'BTC', 1, my_constant_balances[1])
updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'ETH', 1, my_constant_balances[2])
updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'ADA', 1, ada_balance)
updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'ETH2', 1, kraken_balances[0])
updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'SOL', 1, kraken_balances[1])
updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'DOT', 1, kraken_balances[2])
updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'PRE', 1, pre_balance)
# prompt to update for Exodus
showMessage('Claim ALGO Rewards',"Open Exodus \n"
                            "Navigate to Rewards \n"
                            "For ALGO click Get Reward \n"
                            "Claim rewards \n"
                            "After clicking OK, see python window for inputs \n")
atom_balance = float(input("Navigate to ATOM Wallet \n"
                            "Click Send \n"
                             "select ALL, then copy Balance and enter here:  \n"))                            
algo_balance = float(input("Navigate to ALGO Wallet \n"
                            "Click Send \n"
                             "select ALL, then copy Balance and enter here:  \n"))  
updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'ATOM', 1, atom_balance)
updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'ALGO', 1, algo_balance)

driver.execute_script("window.open('https://docs.google.com/spreadsheets/d/1sWJuxtYI-fJ6bUHBWHZTQwcggd30RcOSTMlqIzd1BBo/edit#gid=953264104');")
showMessage("Balances + Review", f'MyConstant: {my_constant_balances[0]} \n' f'Worthy: {worthy_balance} \n' f'Liquid Assets: {liq_assets} \n' f'401k: {HE_balances[2]}')
driver.quit()