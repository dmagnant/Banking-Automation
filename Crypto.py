from Eternl import runEternl
from Exodus import runExodus
from IoPay import runIoPay
from Kraken import runKraken
from Midas import runMidas
from MyConstant import runMyConstant
from Presearch import runPresearch
from Functions import setDirectory, chromeDriverAsUser, showMessage

directory = setDirectory()
driver = chromeDriverAsUser(directory)
driver.implicitly_wait(3)

myConstantBalances = runMyConstant(directory, driver)
adaBalance = runEternl(directory, driver)
krakenBalances = runKraken(directory, driver)
preBalance = runPresearch(directory, driver)
midasBalances = runMidas(directory, driver)
atomBalance = runExodus(directory)
iotxBalance = runIoPay(directory)

driver.execute_script("window.open('https://docs.google.com/spreadsheets/d/1sWJuxtYI-fJ6bUHBWHZTQwcggd30RcOSTMlqIzd1BBo/edit#gid=623829469');")
showMessage("Coin Balances",
f"ALGO:                         {round(krakenBalances[0], 4)} \n"
f"BTC_Midas:                    {round(midasBalances[0], 4)} \n"
f"BTC_MyConstant:           {round(myConstantBalances[1], 4)} \n"
f"ADA:                          {round(adaBalance, 4)} \n"
f"ATOM:                         {round(atomBalance, 4)} \n"
f"ETH_Midas:                    {round(midasBalances[1], 4)} \n"
f"ETH_MyConstant:           {round(myConstantBalances[2], 4)} \n"
f"ETH2:                         {round(krakenBalances[2], 4)} \n"
f"IOTX:                         {round(iotxBalance, 4)} \n"
f"DOT:                          {round(krakenBalances[1], 4)} \n"
f"PRE:                          {round(preBalance[0], 4)} \n"
f"SOL:                          {round(krakenBalances[3], 4)} \n")

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
