from Functions import showMessage, updateSpreadsheet, getCryptocurrencyPrice, setDirectory, updateCryptoPriceInGnucash, updateCoinQuantityFromStakingInGnuCash

def runIoPay(directory):
    showMessage('IOTX balance via IoPay Desktop',"Open IoPay Desktop and Connect Ledger \n"
                                "Connect Ledger > Unlock > Stake your tokens (launches webpage) \n"
                                "My Stakes > Action > Add Stake \n"
                                "stake available balance (may need 100 minimum) \n"
                                "After clicking OK, see python window for inputs \n")
    walletBalance = float(input("copy WALLET BALANCE and paste here:  \n"))
    stakedBalance = float(input("copy TOTAL STAKED AMOUNT and paste here:  \n"))
    iotxBalance = walletBalance + stakedBalance
    updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'IOTX', 1, iotxBalance, "IOTX")
    updateCoinQuantityFromStakingInGnuCash(iotxBalance, 'IOTX')
    iotxPrice = getCryptocurrencyPrice('iotex')['iotex']['usd']
    updateSpreadsheet(directory, 'Asset Allocation', 'Cryptocurrency', 'IOTX', 2, iotxPrice, "IOTX")
    updateCryptoPriceInGnucash('IOTX', format(iotxPrice, ".2f"))

    return iotxBalance

if __name__ == '__main__':
    directory = setDirectory()
    response = runIoPay(directory)
    print('balance: ' + response)