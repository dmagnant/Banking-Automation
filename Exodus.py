from Functions import showMessage

def runExodus():
    showMessage('Claim Rewards',"Open Exodus \n"
                                "Navigate to Rewards \n"
                                "Claim Rewards and restake \n"
                                "After clicking OK, see python window for inputs \n")
    atom_balance = float(input("Navigate to ATOM, find Available balance \n"
                                "copy Balance and enter here:  \n"))                       
    return atom_balance

