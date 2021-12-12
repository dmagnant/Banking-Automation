
from CCVault import runCCVault
from Kraken import runKraken
from MyConstant import runMyConstant
from Exodus import runExodus
from Presearch import runPresearch
from Functions import updateSpreadsheet, setDirectory, chromeDriverAsUser, showMessage

directory = setDirectory()
driver = chromeDriverAsUser(directory)
driver.implicitly_wait(3)

my_constant_balances = runMyConstant(directory, driver)
ada_balance = runCCVault(driver)
kraken_balances = runKraken(directory, driver)
pre_balance = runPresearch(driver)
exodus_balances = runExodus()
showMessage('test', 'test')
updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'ADA', 1, ada_balance)
updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'ALGO', 1, exodus_balances[0])
updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'ATOM', 1, exodus_balances[1])
updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'BTC', 1, my_constant_balances[1])
updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'DOT', 1, kraken_balances[2])
updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'ETH', 1, my_constant_balances[2])
updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'ETH2', 1, kraken_balances[0])
updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'PRE', 1, pre_balance)
updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'SOL', 1, kraken_balances[1])

driver.execute_script("window.open('https://docs.google.com/spreadsheets/d/1sWJuxtYI-fJ6bUHBWHZTQwcggd30RcOSTMlqIzd1BBo/edit#gid=623829469');")

