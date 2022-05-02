from Functions import showMessage

def runIoPay():
    showMessage('Claim Rewards',"Open IoPay Desktop \n"
                                "Navigate to My Stakes \n"
                                "Re-stake available balance (may need 100 minimum) \n"
                                "After clicking OK, see python window for inputs \n")
    iotx_balance = float(input("Navigate to IOTX, find 'wallet balance' (ie staking rewards) \n"
                                "copy wallet balance and enter here (no commas):  \n"))                       
    return iotx_balance