import gspread
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from pykeepass import PyKeePass
from datetime import datetime, timedelta
import time
import os
import pyotp
import psutil
import ctypes
import pygetwindow
import pyautogui
from decimal import Decimal
import csv
import piecash
from piecash import Transaction, Split, GnucashException

def showMessage(header, body): 
    MessageBox = ctypes.windll.user32.MessageBoxW
    MessageBox(None, body, header, 0)

def getOTP(account):
    return pyotp.TOTP(os.environ.get(account)).now()

def getUsername(directory, name):
    keepass_file = directory + r"\Other\KeePass.kdbx"
    KeePass = PyKeePass(keepass_file, password=os.environ.get('KeePass'))
    return KeePass.find_entries(title=name, first=True).username

def getPassword(directory, name):
    keepass_file = directory + r"\Other\KeePass.kdbx"
    KeePass = PyKeePass(keepass_file, password=os.environ.get('KeePass'))
    return KeePass.find_entries(title=name, first=True).password

def checkIfProcessRunning(processName):
    # Iterate over the all the running process
    for proc in psutil.process_iter():
        try:
            # Check if process name contains the given name string.
            if processName.lower() in proc.name().lower():
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass
    return False;

def startExpressVPN():
    os.startfile(r'C:\Program Files (x86)\ExpressVPN\expressvpn-ui\ExpressVPN.exe')
    time.sleep(3)
    EVPN = pygetwindow.getWindowsWithTitle('ExpressVPN')[0]
    EVPN.close()
    # stays open in system tray

def closeExpressVPN():
    if checkIfProcessRunning('ExpressVPN.exe'):
        os.startfile(r'C:\Program Files (x86)\ExpressVPN\expressvpn-ui\ExpressVPN.exe')
        time.sleep(1)
        EVPN = pygetwindow.getWindowsWithTitle('ExpressVPN')[0]
        EVPN.restore()
        EVPN.move(0, 0)
        EVPN.activate()
        pyautogui.leftClick(40, 50)
        time.sleep(1)
        pyautogui.leftClick(40, 280)
        time.sleep(1)

def openGnuCashBook(directory, type, readOnly, openIfLocked):
    if type == 'Finance':
        book = directory + r"\Finances\Personal Finances\Finance.gnucash"
    elif type == 'Home':
        book = directory + r"\Stuff\Home\Finances\Home.gnucash"
    try:
        mybook = piecash.open_book(book, readonly=readOnly, open_if_lock=openIfLocked)
    except GnucashException:
        showMessage("Gnucash file open", f'Close Gnucash file then click OK \n')
        mybook = piecash.open_book(book, readonly=readOnly, open_if_lock=openIfLocked)
    return mybook

def getGnuCashBalance(mybook, account):
    # Get GnuCash Balances
    with mybook as book:
        if account == 'Ally':
            accountpath = "Assets:Ally Checking Account"
        elif account == 'Amex':
            accountpath = "Liabilities:Credit Cards:Amex BlueCash Everyday"
        elif account == 'Barclays':
            accountpath = "Liabilities:Credit Cards:BarclayCard CashForward"
        elif account == 'BoA':
            accountpath = "Liabilities:Credit Cards:BankAmericard Cash Rewards"
        elif account == 'BoA-joint':
            accountpath = "Liabilities:BoA Credit Card"
        elif account == 'Chase':
            accountpath = "Liabilities:Credit Cards:Chase Freedom"
        elif account == 'Discover':
            accountpath = "Liabilities:Credit Cards:Discover It"
        elif account == 'HSA':
            accountpath = "Assets:Non-Liquid Assets:HSA"
        elif account == 'Liquid Assets':
            accountpath = "Assets:Liquid Assets"
        elif account == 'M1':
            accountpath = "Assets:Liquid Assets:M1 Spend"
        elif account == 'MyConstant':
            accountpath = "Assets:Liquid Assets:My Constant"
        elif account == 'TIAA':
            accountpath = "Assets:Liquid Assets:TIAA"
        elif account == 'VanguardPension':
            accountpath = "Assets:Non-Liquid Assets:Pension"
        elif account == 'Worthy':
            accountpath = "Assets:Liquid Assets:Worthy Bonds"
        balance = mybook.accounts(fullname=accountpath).get_balance()
    book.close()
    return balance

def setDirectory():
    return os.environ.get('StorageDirectory')

def chromeDriverAsUser(directory):
    chromedriver = directory + r"\Projects\Coding\webdrivers\chromedriver.exe"
    options = webdriver.ChromeOptions()
    options.add_argument(r"user-data-dir=C:\Users\dmagn\AppData\Local\Google\Chrome\User Data")
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    return webdriver.Chrome(executable_path=chromedriver, options=options)

def updateSpreadsheet(directory, sheetTitle, tabTitle, account, month, value, modified=False):
    json_creds = directory + r"\Projects\Coding\Python\BankingAutomation\Resources\creds.json"
    sheet = gspread.service_account(filename=json_creds).open(sheetTitle)
    worksheet = sheet.worksheet(str(tabTitle))
    cell = getCell(account, month)
    if modified:
        cell = cell.replace(cell[0], chr(ord(cell[0]) + 3))
    worksheet.update(cell, value)

def getCell(account, month):
    if account == 'Amex':
        cellarray = ['K7', 'S7', 'C42', 'K42', 'S42', 'C77', 'K77', 'S77', 'C112', 'K112', 'S112', 'C7']
    elif account == 'Barclays':
        cellarray = ['K10', 'S10', 'C45', 'K45', 'S45', 'C80', 'K80', 'S80', 'C115', 'K115', 'S115', 'C10']
    elif account == 'BoA':
        cellarray = ['K5', 'S5', 'C40', 'K40', 'S40', 'C75', 'K75', 'S75', 'C110', 'K110', 'S110', 'C5']
    elif account == 'BoA-joint':
        cellarray = ['K16', 'S16', 'C52', 'K52', 'S52', 'C88', 'K88', 'S88', 'C124', 'K124', 'S124', 'C16']
    elif account == 'Bonds':
        cellarray = ['F4', 'M4', 'T4', 'F26', 'M26', 'T26', 'F48', 'M48', 'T48', 'F70', 'M70', 'T70']
    elif account == 'Chase':
        cellarray = ['F8', 'S8', 'C43', 'K43', 'S43', 'C78', 'K78', 'S78', 'C113', 'K113', 'S113', 'C8']
    elif account == 'Discover':
        cellarray = ['K6', 'S6', 'C41', 'K41', 'S41', 'C76', 'K76', 'S76', 'C111', 'K111', 'S111', 'C6']
    elif account == 'HE_HSA':
        cellarray = ['B14', 'I14', 'P14', 'B36', 'I36', 'P36', 'B58', 'I58', 'P58', 'B80', 'I80', 'P80']
    elif account == 'Liquid Assets':
        cellarray = ['B4', 'I4', 'P4', 'B26', 'I26', 'P26', 'B48', 'I48', 'P48', 'B70', 'I70', 'P70']
    elif account == 'Vanguard401k':
        cellarray = ['B6', 'I6', 'P6', 'B28', 'I28', 'P28', 'B50', 'I50', 'P50', 'B72', 'I72', 'P72']
    elif account == 'VanguardPension':
        cellarray = ['B8', 'I8', 'P8', 'B30', 'I30', 'P30', 'B52', 'I52', 'P52', 'B74', 'I74', 'P74']
    elif account == 'ADA':
        cellarray = ['I10']
    elif account == 'ALGO':
        cellarray = ['I8']
    elif account == 'ATOM':
        cellarray = ['I11']
    elif account == 'BTC':
        cellarray = ['I9']
    elif account == 'DOT':
        cellarray = ['I15']
    elif account == 'ETH':
        cellarray = ['I12']
    elif account == 'ETH2':
        cellarray = ['I14']
    elif account == 'PRE':
        cellarray = ['I16']
    elif account == 'SOL':
        cellarray = ['I17']
    return cellarray[month - 1]

def getStartAndEndOfPreviousMonth(today, month, year):
    if month == 1:
        startdate = today.replace(month=12, day=1, year=year - 1)
        enddate = today.replace(month=12, day=31, year=year - 1)
    elif month == 3:
        startdate = today.replace(month=2, day=1)
        enddate = today.replace(month=2, day=28)
    elif month in [5, 7, 10, 12]:
        startdate = today.replace(month=month - 1, day=1)
        enddate = today.replace(month=month - 1, day=30)
    else:
        startdate = today.replace(month=month - 1, day=1)
        enddate = today.replace(month=month - 1, day=31)
    return [startdate, enddate]

def getDateRange(today, num_days):
    # Gather last 3 days worth of transactions
    current_date = today.date()    
    date_range = current_date.isoformat()
    day = 1
    while day <= num_days:
        day_before = (current_date - timedelta(days=day)).isoformat()
        date_range = date_range + day_before
        day += 1
    return date_range

def modifyTransactionDescription(description, amount="0.00"):
    if "INTERNET TRANSFER FROM ONLINE SAVINGS ACCOUNT XXXXXX9703" in description.upper():
        description = "Tessa Deposit"
    elif "CITY OF MILWAUKE B2P*MILWWA" in description.upper():
        description = "Water Bill"
    elif "REQUESTED TRANSFER FROM DAN S MAGNANT" in description.upper():
        description = "Dan Deposit"
    elif "DOVENMUEHLE MTG MORTG PYMT" in description.upper():
        description = "Mortgage Payment"
    elif "TRANSFER TO M1 INVEST" in description.upper():
        description = "IRA Transfer"
    elif "INTEREST PAID TO M1 SPEND PLUS" in description.upper():
        description = "Interest paid"
    elif "NORTHWESTERN MUT" in description.upper():
        description = "NM Paycheck"
    elif "PAYPAL" in description.upper() and "10.00" in amount:
        description = "Swagbucks"  
    elif "PAYPAL" in description.upper():
        description = "Paypal"
    elif "LENDING CLUB" in description.upper():
        description = "Lending Club"
    elif "NIELSEN" in description.upper() and "3.00" in amount:
        description = "Pinecone Research"
    elif "VENMO" in description.upper():
        description = "Venmo"
    elif "TIAA" in description.upper():
        description = "TIAA Transfer"
    elif "TRANSFER FROM LINKED BANK" in description.upper():
        description = "TIAA Transfer"
    elif "AMEX EPAYMENT" in description.upper():
        description = "Amex CC"
    elif "COINBASE" in description.upper():
        description = "ADA purchase"
    elif "CHASE CREDIT CRD RWRD" in description.upper():
        description = "Chase CC Rewards"
    elif "CHASE CREDIT CRD AUTOPAY" in description.upper():
        description = "Chase CC"
    elif "DISCOVER E-PAYMENT" in description.upper():
        description = "Discover CC"
    elif "DISCOVER CASH AWARD" in description.upper():
        description = "Discover CC Rewards"
    elif "BARCLAYCARD US CREDITCARD" in description.upper():
        description = "Barclays CC"
    elif "BARCLAYCARD US ACH REWARD" in description.upper():
        description = "Barclays CC Rewards"
    elif "BK OF AMER VISA ONLINE PMT" in description.upper():
        description = "BoA CC"
    elif "ALLY BANK $TRANSFER" in description.upper():
        description = "Ally Transfer"
    return description

def setToAccount(account, row):
    to_account = ''
    row_num = 2 if account in ['BoA', 'BoA-joint', 'Chase', 'Discover'] else 1
    if "BoA CC" in row[row_num]:
        if account == 'Ally':
            to_account = "Liabilities:BoA Credit Card"
        elif account == 'M1':
            to_account = "Liabilities:Credit Cards:BankAmericard Cash Rewards"
    elif "ARCADIA" in row[row_num]:
        to_account = ""
    elif "Tessa Deposit" in row[row_num]:
        to_account = "Tessa's Contributions"
    elif "Water Bill" in row[row_num]:
        to_account = "Expenses:Utilities:Water"
    elif "Dan Deposit" in row[row_num]:
        to_account = "Dan's Contributions"
    elif "Mortgage Payment" in row[row_num]:
        to_account = "Liabilities:Mortgage Loan"
    elif "Swagbucks" in row[row_num]:
        to_account = "Income:Market Research"
    elif "NM Paycheck" in row[row_num]:
        to_account = "Income:Salary"
    elif "GOOGLE FI" in row[row_num].upper():
        to_account = "Expenses:Utilities:Phone"
    elif "TIAA Transfer" in row[row_num]:
        to_account = "Assets:Liquid Assets:TIAA"
    elif "ADA PURCHASE" in row[row_num].upper():
        to_account = "Assets:Non-Liquid Assets:CryptoCurrency"
    elif "Pinecone Research" in row[row_num]:
        to_account = "Income:Market Research"
    elif "IRA Transfer" in row[row_num]:
        to_account = "Assets:Non-Liquid Assets:Roth IRA"
    elif "Lending Club" in row[row_num]:
        to_account = "Assets:Non-Liquid Assets:MicroLoans"
    elif "CHASE CC REWARDS" or "DISCOVER CC REWARDS" or "BARCLAYS CC REWARDS" or "REDEMPTION CREDIT" or "CASH REWARD" in row[row_num].upper():
        to_account = "Income:Credit Card Rewards"
    elif "Chase CC" in row[row_num]:
        to_account = "Liabilities:Credit Cards:Chase Freedom"
    elif "Discover CC" in row[row_num]:
        to_account = "Liabilities:Credit Cards:Discover It"
    elif "Amex CC" in row[row_num]:
        to_account = "Liabilities:Credit Cards:Amex BlueCash Everyday"
    elif "Barclays CC" in row[row_num]:
        to_account = "Liabilities:Credit Cards:BarclayCard CashForward"
    elif "Ally Transfer" in row[row_num]:
        to_account = "Expenses:Joint Expenses"
    elif "BP#" in row[row_num]:
        to_account = "Expenses:Transportation:Gas (Vehicle)"
    elif "CAT DOCTOR" in row[row_num]:
        to_account = "Expenses:Medical:Vet"
    elif "PARKING" in row[row_num]:
        to_account = "Expenses:Transportation:Parking"
    elif "PROGRESSIVE" in row[row_num]:
        to_account = "Expenses:Transportation:Car Insurance"
    elif "SPECTRUM" in row[row_num].upper():
            to_account = "Expenses:Utilities:Internet"     
    elif "UBER" in row[row_num].upper():
        to_account = "Expenses:Travel:Ride Services" if account in ['BoA-joint', 'Ally'] else "Expenses:Transportation:Ride Services"
    elif "TECH WAY AUTO SERV" in row[row_num].upper():
        to_account = "Expenses:Transportation:Car Maintenance"
    elif "INTEREST PAID" in row[row_num].upper():
        to_account = "Income:Interest" if account in ['BoA-joint', 'Ally'] else "Income:Investments:Interest"

    if not to_account:
        if (row[row_num].upper() in ['HOMEDEPOT.COM', 'THE HOME DEPOT']):
            if account in ['BoA-joint', 'Ally']:
                to_account = "Expenses:Home Depot"
    
    if not to_account:
        if (row[row_num].upper() in ['AMAZON', 'AMZN']):
            to_account = "Expenses:Amazon"

    if not to_account:
        if account == "Chase" or "Discover":
            if row[3] or row[4] == "Groceries" or "Supermarkets":
                to_account = "Expenses:Groceries"
        if not to_account:
            if (row[row_num].upper() in ['PICK N SAVE', 'KOPPA', 'KETTLE RANGE', 'WHOLE FOODS', 'WHOLEFDS', 'TARGET']):
                to_account = "Expenses:Groceries"

    if not to_account:
        if account == "Chase" or "Discover":
            if row[3] or row[4] == "Food & Drink" or "Restaurants":
                to_account = "Expenses:Bars & Restaurants"
        if not to_account:
            if (row[row_num].upper() in ['MCDONALD', 'GRUBHUB', 'JIMMY JOHN', 'COLECTIVO']):
                to_account = "Expenses:Bars & Restaurants"
    
    if not to_account:
            to_account = "Expenses:Other"
    return to_account

def formatTransactionVariables(account, row):
    cc_payment = False
    description = row[1]
    if account == 'Ally':
        postdate = datetime.strptime(row[0], '%Y-%m-%d')
        description = row[1]
        amount = Decimal(row[2])
        from_account = "Assets:Ally Checking Account"
        review_trans_path = row[0] + ", " + row[1] + ", " + row[2] + "\n"
    elif account == 'Amex':
        postdate = datetime.strptime(row[0], '%m/%d/%Y')
        description = row[1]
        amount = -Decimal(row[2])
        if "AUTOPAY PAYMENT" in row[1]:
            cc_payment = True
        from_account = "Liabilities:Credit Cards:Amex BlueCash Everyday"
        review_trans_path = row[0] + ", " + row[1] + ", " + row[2] + "\n"
    elif account == 'Barclays':
        postdate = datetime.strptime(row[0], '%m/%d/%Y')
        description = row[1]
        amount = Decimal(row[3])
        if "Payment Received" in row[1]:
            cc_payment = True
        from_account = "Liabilities:Credit Cards:BarclayCard CashForward"
        review_trans_path = row[0] + ", " + row[1] + ", " + row[3] + "\n"
    elif account == 'BoA':
        postdate = datetime.strptime(row[0], '%m/%d/%Y')
        description = row[2]
        amount = Decimal(row[4])
        if "BA ELECTRONIC PAYMENT" in row[2]:
            cc_payment = True
        from_account = "Liabilities:Credit Cards:BankAmericard Cash Rewards"
        review_trans_path = row[0] + ", " + row[2] + ", " + row[4] + "\n"
    elif account == 'BoA-joint':
        postdate = datetime.strptime(row[0], '%m/%d/%Y')
        description = row[2]
        amount = Decimal(row[4])
        if "BA ELECTRONIC PAYMENT" in row[2]:
            cc_payment = True
        from_account = "Liabilities:BoA Credit Card"
        review_trans_path = row[0] + ", " + row[2] + ", " + row[4] + "\n"
    elif account == 'Chase':
        postdate = datetime.strptime(row[1], '%m/%d/%Y')
        description = row[2]
        amount = Decimal(row[5])
        if "AUTOMATIC PAYMENT" in row[2]:
            cc_payment = True
        from_account = "Liabilities:Credit Cards:Chase Freedom"
        review_trans_path = row[1] + ", " + row[2] + ", " + row[5] + "\n"
    elif account == 'Discover':
        postdate = datetime.strptime(row[1], '%m/%d/%Y')
        description = row[2]
        amount = -Decimal(row[3])
        if "DIRECTPAY FULL BALANCE" in row[2]:
            cc_payment = True
        from_account = "Liabilities:Credit Cards:Discover It"
        review_trans_path = row[1] + ", " + row[2] + ", " + row[3] + "\n"
    elif account == 'M1':
        postdate = datetime.strptime(row[0], '%Y-%m-%d')
        description = row[1]
        amount = Decimal(row[2])
        from_account = "Assets:Liquid Assets:M1 Spend"
        review_trans_path = row[0] + ", " + row[1] + ", " + row[2] + "\n"
    return [postdate, description, amount, cc_payment, from_account, review_trans_path]

def compileGnuTransactions(account, transactions_csv, gnu_csv, mybook, driver, directory, date_range, line_start=1):
    import_csv = directory + r"\Projects\Coding\Python\BankingAutomation\Resources\import.csv"
    open(import_csv, 'w', newline='').truncate()
    if account == 'Ally':
        gnu_account = "Assets:Ally Checking Account"
    elif account == 'M1':
        gnu_account = "Assets:Liquid Assets:M1 Spend"
    
    # retrieve transactions from GnuCash
    transactions = [tr for tr in mybook.transactions
                    if str(tr.post_date.strftime('%Y-%m-%d')) in date_range
                    for spl in tr.splits
                    if spl.account.fullname == gnu_account]
    for tr in transactions:
        date = str(tr.post_date.strftime('%Y-%m-%d'))
        description = str(tr.description)
        for spl in tr.splits:
            amount = format(spl.value, ".2f")
            if spl.account.fullname == gnu_account:
                row = date, description, str(amount)
                csv.writer(open(gnu_csv, 'a', newline='')).writerow(row)

    for row in csv.reader(open(transactions_csv, 'r'), delimiter=','):
        if row not in csv.reader(open(gnu_csv, 'r'), delimiter=','):
            csv.writer(open(import_csv, 'a', newline='')).writerow(row)
    review_trans = importGnuTransaction(account, import_csv, mybook, driver, directory, line_start)
    return review_trans

def importGnuTransaction(account, transactions_csv, mybook, driver, directory, line_start=1):
    review_trans = ''
    row_count = 0
    line_count = 0
    energy_bill_num = 0
    for row in csv.reader(open(transactions_csv), delimiter=','):
        row_count += 1
        # skip header line
        if line_count < line_start:
            line_count += 1
        else:
            transaction_variables = formatTransactionVariables(account, row)
            # Skip credit card payments from CC bills (already captured through Checking accounts)
            if transaction_variables[3]:
                continue
            else:
                description = transaction_variables[1]
                postdate = transaction_variables[0]
                from_account = transaction_variables[4]
                amount = transaction_variables[2]
                to_account = setToAccount(account, row)
                if 'ARCADIA' in description.upper():
                    energy_bill_num += 1
                    amount = getEnergyBillAmounts(driver, directory, transaction_variables[2], energy_bill_num)
                elif 'NM PAYCHECK' or "ADA PURCHASE" in description.upper():
                    review_trans = review_trans + transaction_variables[5]
                else:
                    if to_account == "Expenses:Other":
                        review_trans = review_trans + transaction_variables[5]
                writeGnuTransaction(mybook, description, postdate, amount, from_account, to_account)
    return review_trans

def writeGnuTransaction(mybook, description, postdate, amount, from_account, to_account=''):
    with mybook as book:
        if "Contribution + Interest" in description:
            split = [Split(value=amount[0], memo="scripted", account=mybook.accounts(fullname="Income:Investments:Interest")),
                    Split(value=amount[1], memo="scripted",account=mybook.accounts(fullname="Income:Employer Pension Contributions")),
                    Split(value=amount[2], memo="scripted",account=mybook.accounts(fullname=from_account))]
        elif "HSA Statement" in description:
            split = [Split(value=amount[0], account=mybook.accounts(fullname=to_account)),
                    Split(value=amount[1], account=mybook.accounts(fullname=from_account[0])),
                    Split(value=amount[2], account=mybook.accounts(fullname=from_account[1]))]
        elif "ARCADIA" in description:
            split=[Split(value=amount[0], memo="Arcadia Membership Fee", account=mybook.accounts(fullname="Expenses:Utilities:Arcadia Membership")),
                    Split(value=amount[1], memo="Solar Rebate", account=mybook.accounts(fullname="Expenses:Utilities:Arcadia Membership")),
                    Split(value=amount[2], account=mybook.accounts(fullname="Expenses:Utilities:Electricity")),
                    Split(value=amount[3], account=mybook.accounts(fullname="Expenses:Utilities:Gas")),
                    Split(value=amount[4], account=mybook.accounts(fullname=from_account))]
        elif "NM Paycheck" in description:
            split = [Split(value=round(Decimal(1621.40), 2), memo="scripted",account=mybook.accounts(fullname=from_account)),
                    Split(value=round(Decimal(173.36), 2), memo="scripted",account=mybook.accounts(fullname="Assets:Non-Liquid Assets:401k")),
                    Split(value=round(Decimal(250.00), 2), memo="scripted",account=mybook.accounts(fullname="Assets:Liquid Assets:Promos")),
                    Split(value=round(Decimal(5.49), 2), memo="scripted",account=mybook.accounts(fullname="Expenses:Medical:Dental")),
                    Split(value=round(Decimal(36.22), 2), memo="scripted",account=mybook.accounts(fullname="Expenses:Medical:Health")),
                    Split(value=round(Decimal(2.67), 2), memo="scripted",account=mybook.accounts(fullname="Expenses:Medical:Vision")),
                    Split(value=round(Decimal(168.54), 2), memo="scripted",account=mybook.accounts(fullname="Expenses:Income Taxes:Social Security")),
                    Split(value=round(Decimal(39.42), 2), memo="scripted",account=mybook.accounts(fullname="Expenses:Income Taxes:Medicare")),
                    Split(value=round(Decimal(305.08), 2), memo="scripted",account=mybook.accounts(fullname="Expenses:Income Taxes:Federal Tax")),
                    Split(value=round(Decimal(157.03), 2), memo="scripted",account=mybook.accounts(fullname="Expenses:Income Taxes:State Tax")),
                    Split(value=round(Decimal(130.00), 2), memo="scripted",account=mybook.accounts(fullname="Assets:Non-Liquid Assets:HSA")),
                    Split(value=-round(Decimal(2889.21), 2), memo="scripted",account=mybook.accounts(fullname=to_account))]
        # elif "Barclays CC Rewards" in description:
        #     split = [Split(value=amount, memo="scripted", account=mybook.accounts(fullname=to_account)),
        #             Split(value=-amount, memo="scripted", account=mybook.accounts(fullname=from_account))]
        else:
            split = [Split(value=-amount, memo="scripted", account=mybook.accounts(fullname=to_account)),
                    Split(value=amount, memo="scripted", account=mybook.accounts(fullname=from_account))]
        Transaction(post_date=postdate.date(), currency=mybook.currencies(mnemonic="USD"), description=description, splits=split)
        book.save()
        book.flush()
    book.close()

def getEnergyBillAmounts(driver, directory, amount, energy_bill_num):
    if energy_bill_num == 1:
        # Get balances from Arcadia
        driver.execute_script("window.open('https://login.arcadia.com/email');")
        driver.implicitly_wait(5)
        # switch to last window
        driver.switch_to.window(driver.window_handles[len(driver.window_handles)-1])
        # get around bot-prevention by logging in twice
        num = 1
        while num <3:
            try:
                # click Sign in with email
                driver.find_element_by_xpath("/html/body/div/main/div[1]/div/div/div[1]/div/a").click()
                time.sleep(1)
            except NoSuchElementException:
                exception = "sign in page loaded already"
            try:
                # Login
                driver.find_element_by_xpath("/html/body/div[1]/main/div[1]/div/form/div[1]/div[1]/input").send_keys(getUsername(directory, 'Arcadia Power'))
                time.sleep(1)
                driver.find_element_by_xpath("/html/body/div[1]/main/div[1]/div/form/div[1]/div[2]/input").send_keys(getPassword(directory, 'Arcadia Power'))
                time.sleep(1)
                driver.find_element_by_xpath("/html/body/div[1]/main/div[1]/div/form/div[2]/button").click()
                time.sleep(1)
                # Get Billing page
                driver.get("https://home.arcadia.com/billing")
            except NoSuchElementException:
                exception = "already signed in"
            num += 1
            if not driver.find_element_by_xpath('/html/body/div[1]/div[2]/div/div[1]/h1'):
                showMessage("Login Check", 'Confirm Login to Arcadia, (manually if necessary) \n' 'Then click OK \n')
    else:
        # switch to last window
        driver.switch_to.window(driver.window_handles[len(driver.window_handles)-1])
        driver.get("https://home.arcadia.com/billing")
    statement_row = 1
    statement_found = "no"                     
    while statement_found == "no":
        # Capture statement balance
        arcadia_balance = driver.find_element_by_xpath("/html/body/div[1]/div[2]/div/div[2]/div[2]/div/div[2]/ul/li[" + str(statement_row) + "]/div/div/p").text.strip('$')
        formatted_amount = "{:.2f}".format(abs(amount))
        if arcadia_balance == formatted_amount:
            # click to view statement
            driver.find_element_by_xpath("/html/body/div[1]/div[2]/div/div[2]/div[2]/div/div[2]/ul/li[" + str(statement_row) + "]/div/div/p").click()
            statement_found = "yes"
        else:
            statement_row += 1
    # comb through lines of Arcadia Statement for Arcadia Membership (and Free trial rebate), Community Solar lines (3)
    arcadia_statement_lines_left = True
    statement_row = 1
    solar = 0
    arcadia_membership = 0
    while arcadia_statement_lines_left:
        try:
            # read the header to get transaction description
            statement_trans = driver.find_element_by_xpath("/html/body/div[1]/div[2]/div[2]/div[5]/ul/li[" + str(statement_row) + "]/div/h2").text
            if statement_trans == "Arcadia Membership":
                arcadia_membership = Decimal(driver.find_element_by_xpath("/html/body/div[1]/div[2]/div[2]/div[5]/ul/li[" + str(statement_row) + "]/div/p").text.strip('$'))
                arcadiaamt = Decimal(arcadia_membership)
            elif statement_trans == "Free Trial":
                arcadia_membership = arcadia_membership + Decimal(driver.find_element_by_xpath("/html/body/div[1]/div[2]/div[2]/div[5]/ul/li[" + str(statement_row) + "]/div/p").text.strip('$'))
            elif statement_trans == "Community Solar":
                solar = solar + Decimal(driver.find_element_by_xpath("/html/body/div[1]/div[2]/div[2]/div[5]/ul/li[" + str(statement_row) + "]/div/p").text.strip('$'))
            elif statement_trans == "WE Energies Utility":
                we_bill = driver.find_element_by_xpath("/html/body/div[1]/div[2]/div[2]/div[5]/ul/li[" + str(statement_row) + "]/div/p").text
            statement_row += 1
        except NoSuchElementException:
            arcadia_statement_lines_left = False
    arcadiaamt = Decimal(arcadia_membership)
    solaramt = Decimal(solar)
    # Get balances from WE Energies
    if energy_bill_num == 1:
        driver.execute_script("window.open('https://www.we-energies.com/secure/auth/l/acct/summary_accounts.aspx');")
        # switch to last window
        driver.switch_to.window(driver.window_handles[len(driver.window_handles)-1])
        try:
            ## LOGIN
            driver.find_element_by_xpath("//*[@id='signInName']").send_keys(getUsername(directory, 'WE-Energies (Home)'))
            driver.find_element_by_xpath("//*[@id='password']").send_keys(getPassword(directory, 'WE-Energies (Home)'))
            # click Login
            driver.find_element_by_xpath("//*[@id='next']").click()
            time.sleep(4)
            # close out of app notice
            driver.find_element_by_xpath("//*[@id='notInterested']/a").click
        except NoSuchElementException:
            exception = "caught"
        # Click View bill history
        driver.find_element_by_xpath("//*[@id='mainContentCopyInner']/ul/li[2]/a").click()
        time.sleep(4)
    bill_row = 2
    bill_column = 7
    bill_found = "no"
    # find bill based on comparing amount from Arcadia (we_bill)
    while bill_found == "no":
        # capture date
        we_bill_path = "/html/body/div[1]/div[1]/form/div[5]/div/div/div/div/div[6]/div[2]/div[2]/div/table/tbody/tr[" + str(bill_row) + "]/td[" + str(bill_column) + "]/span/span"
        we_bill_amount = driver.find_element_by_xpath(we_bill_path).text
        if we_bill == we_bill_amount:
            bill_found = "yes"
        else:
            bill_row += 1
    # capture gas charges
    bill_column -= 2
    we_amt_path = "/html/body/div[1]/div[1]/form/div[5]/div/div/div/div/div[6]/div[2]/div[2]/div/table/tbody/tr[" + str(bill_row) + "]/td[" + str(bill_column) + "]/span"
    gasamt = Decimal(driver.find_element_by_xpath(we_amt_path).text.strip('$'))
    # capture electricity charges
    bill_column -= 2
    we_amt_path = "/html/body/div[1]/div[1]/form/div[5]/div/div/div/div/div[6]/div[2]/div[2]/div/table/tbody/tr[" + str(bill_row) + "]/td[" + str(bill_column) + "]/span"
    electricityamt = Decimal(driver.find_element_by_xpath(we_amt_path).text.strip('$'))
    return [arcadiaamt, solaramt, electricityamt, gasamt, amount]