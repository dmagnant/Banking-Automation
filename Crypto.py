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
atom_balance = runExodus()

updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'ADA', 1, ada_balance)
updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'ALGO', 1, kraken_balances[3])
updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'ATOM', 1, atom_balance)
updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'BTC', 1, my_constant_balances[1])
updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'DOT', 1, kraken_balances[2])
updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'ETH', 1, my_constant_balances[2])
updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'ETH2', 1, kraken_balances[0])
updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'PRE', 1, pre_balance[0])
updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'SOL', 1, kraken_balances[1])

driver.execute_script("window.open('https://docs.google.com/spreadsheets/d/1sWJuxtYI-fJ6bUHBWHZTQwcggd30RcOSTMlqIzd1BBo/edit#gid=623829469');")

# # write cardano transaction from coinbase
# mybook = openGnuCashBook(directory, 'Finance', False, False)
# from_account = 'Assets:Liquid Assets:M1 Spend'
# to_account = 'Assets:Non-Liquid Assets:CryptoCurrency:Cardano'
# fee_account = 'Expenses:Bank Fees:Coinbase Fee'
# amount = Decimal(50.00)
# description = 'ADA purchase'
# today = datetime.today()
# year = today.year
# postdate = today.replace(month=1, day=1, year=year)
# with mybook as book:
#     split = [Split(value=-amount, memo="scripted", account=mybook.accounts(fullname=from_account)),
#             Split(value=round(amount-Decimal(1.99), 2), quantity=round(Decimal(35.052832), 6), memo="scripted", account=mybook.accounts(fullname=to_account)),
#             Split(value=round(Decimal(1.99),2), memo="scripted", account=mybook.accounts(fullname=fee_account))]
#     Transaction(post_date=postdate.date(), currency=mybook.currencies(mnemonic="USD"), description=description, splits=split)
#     book.save()
#     book.flush()
# book.close()

## update price in gnu cash
# with mybook as book:
#     p = Price(commodity=mybook.commodities(mnemonic="ADA"), currency=mybook.currencies(mnemonic="USD"), date=today.date(), value=Decimal('1.99'))
#     book.save()
#     book.flush()
# book.close()

# get dollars invested balance (must be run per coin)
# mybook = openGnuCashBook(directory, 'Finance', True, True)
# gnu_account = "Assets:Non-Liquid Assets:CryptoCurrency"
# total = 0
# # retrieve transactions from GnuCash
# transactions = [tr for tr in mybook.transactions
#                 for spl in tr.splits
#                 if spl.account.fullname == gnu_account]
# for tr in transactions:
#     for spl in tr.splits:
#         amount = format(spl.value, ".2f")
#         if spl.account.fullname == gnu_account:
#             total += abs(float(amount))
# print(total)

# # get total investment (dollars) # # 
# def sumDollarInvestment(mybook, gnu_account):
#     sum = 0
#     # retrieve transactions from GnuCash
#     transactions = [tr for tr in mybook.transactions
#                     for spl in tr.splits
#                     if spl.account.fullname == gnu_account]
#     for tr in transactions:
#         for spl in tr.splits:
#             amount = format(spl.value, ".2f")
#             if spl.account.fullname == gnu_account:
#                 sum += abs(float(amount))
#     return sum

# def getCryptoInvestmentInDollars():
#     mybook = openGnuCashBook(directory, 'Finance', True, True)
#     account = "Assets:Non-Liquid Assets:CryptoCurrency"
#     total = 0
#     for i in mybook.accounts(fullname=account).children:
#         if len(i.children) > 1:
#             for j in i.children:
#                 gnu_account = "Assets:Non-Liquid Assets:CryptoCurrency:" + i.name + ":" + j.name
#                 total += sumDollarInvestment(mybook, gnu_account)
#         else:
#             gnu_account = "Assets:Non-Liquid Assets:CryptoCurrency:" + i.name
#             total += sumDollarInvestment(mybook, gnu_account)
#     print('total: ', total)
#     return total
