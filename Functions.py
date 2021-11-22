import gspread
from selenium import webdriver
from pykeepass import PyKeePass
import os
import psutil
import time
import ctypes
import pygetwindow
from datetime import timedelta
import pyautogui
import gspread
import piecash
from piecash import GnucashException

def showMessage(header, body): 
    MessageBox = ctypes.windll.user32.MessageBoxW
    MessageBox(None, body, header, 0)

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
        if account == 'M1':
            gnuCashAccount = mybook.accounts(fullname="Assets:Liquid Assets:M1 Spend")
        elif account == 'TIAA':
            gnuCashAccount = mybook.accounts(fullname="Assets:Liquid Assets:TIAA")
        elif account == 'Ally':
            gnuCashAccount = mybook.accounts(fullname="Assets:Ally Checking Account")
        elif account == 'Amex':
            gnuCashAccount = mybook.accounts(fullname="Liabilities:Credit Cards:Amex BlueCash Everyday")
        elif account == 'Barclays':
            gnuCashAccount = mybook.accounts(fullname="Liabilities:Credit Cards:BarclayCard CashForward")
        elif account == 'BoA':
            gnuCashAccount = mybook.accounts(fullname="Liabilities:Credit Cards:BankAmericard Cash Rewards")
        elif account == 'BoA-joint':
            gnuCashAccount = mybook.accounts(fullname="Liabilities:BoA Credit Card")
        elif account == 'Chase':
            gnuCashAccount = mybook.accounts(fullname="Liabilities:Credit Cards:Chase Freedom")
        elif account == 'Discover':
            gnuCashAccount = mybook.accounts(fullname="Liabilities:Credit Cards:Discover It")
        elif account == 'VanguardPension':
            gnuCashAccount = mybook.accounts(fullname="Assets:Non-Liquid Assets:Pension")
        balance = gnuCashAccount.get_balance()
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

def updateSpreadsheet(directory, type, year, account, month, value, modified=False):
    json_creds = directory + r"\Projects\Coding\Python\BankingAutomation\Resources\creds.json"
    sheet = gspread.service_account(filename=json_creds).open(type)
    if type == 'Home':
        worksheet = sheet.worksheet(str(year) + " Balance")
    else: 
        worksheet = sheet.worksheet(str(year))
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
        description = "Barclays CC rewards"
    elif "BK OF AMER VISA ONLINE PMT" in description.upper():
        description = "BoA CC"
    elif "ALLY BANK $TRANSFER" in description.upper():
        description = "Ally Transfer"
    return description

def setToAccount(account, row):
    to_account = ''
    if account in ['BoA', 'BoA-joint', 'Chase', 'Discover']:
        row_num = 2
    else:
        row_num = 1
    if "BoA CC" in row[row_num]:
        if account == 'Ally':
            to_account = "Liabilities:BoA Credit Card"
        elif account == 'M1':
            to_account = "Liabilities:Credit Cards:BankAmericard Cash Rewards"
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
    elif "COINBASE" in row[row_num].upper():
        to_account = "Assets:Non-Liquid Assets:CryptoCurrency"
    elif "Pinecone Research" in row[row_num]:
        to_account = "Income:Market Research"
    elif "IRA Transfer" in row[row_num]:
        to_account = "Assets:Non-Liquid Assets:Roth IRA"
    elif "Lending Club" in row[row_num]:
        to_account = "Assets:Non-Liquid Assets:MicroLoans"
    elif "Chase CC Rewards" in row[row_num]:
        to_account = "Income:Credit Card Rewards"
    elif "Chase CC" in row[row_num]:
        to_account = "Liabilities:Credit Cards:Chase Freedom"
    elif "Discover CC Rewards" in row[row_num]:
        to_account = "Income:Credit Card Rewards"
    elif "Discover CC" in row[row_num]:
        to_account = "Liabilities:Credit Cards:Discover It"
    elif "Amex CC" in row[row_num]:
        to_account = "Liabilities:Credit Cards:Amex BlueCash Everyday"
    elif "Barclays CC Rewards" in row[row_num]:
        to_account = "Income:Credit Card Rewards"
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
        if account in ['BoA-joint', 'Ally']:
            to_account = "Expenses:Travel:Ride Services"
        else:
            to_account = "Expenses:Transportation:Ride Services"
    elif "INTEREST PAID" in row[row_num].upper():
        if account in ['BoA-joint', 'Ally']:
            to_account = "Income:Interest"
        else:
            to_account = "Income:Investments:Interest"

    if not to_account:
        for i in ['REDEMPTION CREDIT', 'CASH REWARD']:
            if i in row[row_num].upper():
                to_account = "Income:Credit Card Rewards"
    
    if not to_account:
        for i in ['HOMEDEPOT.COM', 'THE HOME DEPOT']:
            if i in row[row_num].upper():
                if account in ['BoA-joint', 'Ally']:
                    to_account = "Expenses:Home Depot"
    
    if not to_account:
        for i in ['AMAZON', 'AMZN']:
            if i in row[row_num].upper():
                to_account = "Expenses:Amazon"

    if not to_account:
        if account == "Chase":
            if row[3] == "Groceries":
                to_account = "Expenses:Groceries"
        elif account == "Discover":
            if row[4] == "Groceries":
                to_account = "Expenses:Groceries"
        if not to_account:
            for i in ['PICK N SAVE', 'KOPPA', 'KETTLE RANGE', 'WHOLE FOODS', 'WHOLEFDS', 'TARGET']:
                if i in row[row_num].upper():
                    to_account = "Expenses:Groceries"

    if not to_account:
        if account == "Chase":
            if row[3] == "Food & Drink":
                to_account = "Expenses:Bars & Restaurants"
        elif account == "Discover":
            if row[4] == "Restaurants":
                to_account = "Expenses:Bars & Restaurants"
        if not to_account:
            for i in ['MCDONALD', 'GRUBHUB', 'JIMMY JOHN', 'COLECTIVO']:
                if i in row[row_num].upper():
                    to_account = "Expenses:Bars & Restaurants"
    
    if not to_account:
            to_account = "Expenses:Other"
    return to_account