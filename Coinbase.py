from datetime import datetime, timedelta
from decimal import Decimal

import gspread
from piecash import GnucashException, Price, Split, Transaction
from pycoingecko import CoinGeckoAPI

from Eternl import runEternl
from Exodus import runExodus
from Functions import (chromeDriverAsUser, getAccountPath,
                       getCryptocurrencyPrice, getGnuCashBalance,
                       openGnuCashBook, setDirectory, showMessage,
                       updateCryptoPrices, updateSpreadsheet)
from IoPay import runIoPay
from Kraken import runKraken
from Midas import runMidas
from MyConstant import runMyConstant
from Presearch import runPresearch

# from Crypto import getCryptoInvestmentInDollars

# # write cardano transaction from coinbase
# directory = setDirectory()
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
#                 if float(amount) > 0:
#                     sum += float(amount)
#     return sum

# def getCryptoInvestmentInDollars(directory=setDirectory()):
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

# getCryptoInvestmentInDollars()
