from Eternl import runEternl
from Exodus import runExodus
from IoPay import runIoPay
from Kraken import runKraken
from Midas import runMidas
from MyConstant import runMyConstant
from Presearch import runPresearch
from datetime import datetime, timedelta
from decimal import Decimal
from piecash import Transaction, Split, GnucashException, Price

from Functions import setDirectory, chromeDriverAsUser, showMessage, openGnuCashBook, getCryptocurrencyPrice, updateCryptoPrice, getAccountPath

def runCrypto(directory, driver):
    # directory = setDirectory()
    # driver = chromeDriverAsUser(directory)
    driver.implicitly_wait(3)
    driver.get("https://docs.google.com/spreadsheets/d/1sWJuxtYI-fJ6bUHBWHZTQwcggd30RcOSTMlqIzd1BBo/edit#gid=623829469")

    myConstantBalances = runMyConstant(directory, driver)
    adaBalance = runEternl(directory, driver)
    krakenBalances = runKraken(directory, driver)
    preBalance = runPresearch(directory, driver)
    midasBalances = runMidas(directory, driver)
    atomBalance = runExodus(directory)
    iotxBalance = runIoPay(directory)

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

if __name__ == '__main__':
    directory = setDirectory()
    driver = chromeDriverAsUser()
    response = runCrypto(directory, driver)
