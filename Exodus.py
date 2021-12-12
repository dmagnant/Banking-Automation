from Functions import showMessage

def runExodus():
    showMessage('Claim ALGO Rewards',"Open Exodus \n"
                                "Navigate to Rewards \n"
                                "Claim Rewards for ALGO and ATOM \n"
                                "restake ATOM rewards \n"
                                "After clicking OK, see python window for inputs \n")
    algo_balance = float(input("Navigate to ALGO, find Available balance \n"
                                "copy Balance and enter here:  \n"))
    atom_balance = float(input("Navigate to ATOM, find Available balance \n"
                                "copy Balance and enter here:  \n"))                       
    return [algo_balance, atom_balance]

