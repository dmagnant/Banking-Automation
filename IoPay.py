from Functions import showMessage

def runIoPay():
    showMessage('Claim Rewards',"Open IoPay Desktop \n"
                                "Navigate to Rewards \n"
                                "Claim Rewards and restake \n"
                                "After clicking OK, see python window for inputs \n")
    iotx_balance = float(input("Navigate to IOTX, find Available balance \n"
                                "copy Balance and enter here:  \n"))                       
    return iotx_balance