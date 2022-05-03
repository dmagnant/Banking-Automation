from Functions import showMessage, updateSpreadsheet, getCryptocurrencyPrice

def runIoPay(directory):
    showMessage('IOTX balance via IoPay Desktop',"Open IoPay Desktop and Connect Ledger \n"
                                "Connect Ledger > Unlock > Stake your tokens (launches webpage) \n"
                                "My Stakes > Action > Add Stake \n"
                                "stake available balance (may need 100 minimum) \n"
                                "After clicking OK, see python window for inputs \n")
    wallet_balance = float(input("copy WALLET BALANCE and paste here:  \n"))
    staked_balance = float(input("copy TOTAL STAKED AMOUNT and paste here:  \n"))
    iotxBalance = wallet_balance + staked_balance
    updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'IOTX', 1, iotxBalance, "IOTX")
    iotxPrice = getCryptocurrencyPrice('iotex')['iotex']['usd']
    updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'IOTX', 2, iotxPrice, "IOTX")
    return iotxBalance
