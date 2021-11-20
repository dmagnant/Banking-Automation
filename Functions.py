import gspread
from selenium import webdriver
import socket
from pykeepass import PyKeePass
import os
import psutil
import time
import ctypes
import pygetwindow
import pyautogui
import gspread
import piecash
from piecash import GnucashException

def showMessage(header, body): 
    MessageBox = ctypes.windll.user32.MessageBoxW
    MessageBox(None, body, header, 0)

def setDirectory():
    #get computer name
    computer = socket.gethostname()
    # computer-specific file paths
    if computer == "Big-Bertha":
        return r"C:\Users\dmagn\Google Drive"
    elif computer == "Black-Betty":
        return r"D:\Google Drive"

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
    cell = getCell(month, account)
    if modified:
        cell = cell.replace(cell[0], chr(ord(cell[0]) + 3))
    worksheet.update(cell, value)

def getCell(month, account):
    if month == 1:
        if account == 'Amex':
            cell = 'K7'
        elif account == 'Barclays':
            cell = 'K10'        
        elif account == 'BoA':
            cell = 'K5'
        elif account == 'BoA-joint':
            cell = 'K16'
        elif account == 'Bonds':
            cell = 'F4'
        elif account == 'Chase':
            cell = 'F8'
        elif account == 'Discover':
            cell = 'K6'  
        elif account == 'HE_HSA':
            cell = 'B14'
        elif account == 'Liquid Assets':
            cell = 'B4'
        elif account == 'Vanguard401k':
            cell = 'B6'
        elif account == 'VanguardPension':
            cell = 'B8'
    elif month == 2:
        if account == 'Amex':
            cell = 'S7'
        elif account == 'Barclays':
            cell = 'S10'        
        elif account == 'BoA':
            cell = 'S5'
        elif account == 'BoA-joint':
            cell = 'S16'
        elif account == 'Bonds':
            cell = 'M4'
        elif account == 'Chase':
            cell = 'S8'
        elif account == 'Discover':
            cell = 'S6'  
        elif account == 'HE_HSA':
            cell = 'I14'
        elif account == 'Liquid Assets':
            cell = 'I4'
        elif account == 'Vanguard401k':
            cell = 'I6'
        elif account == 'VanguardPension':
            cell = 'I8'
    elif month == 3:
        if account == 'Amex':
            cell = 'C42'
        elif account == 'Barclays':
            cell = 'C45'        
        elif account == 'BoA':
            cell = 'C40'
        elif account == 'BoA-joint':
            cell = 'C52'
        elif account == 'Bonds':
            cell = 'T4'
        elif account == 'Chase':
            cell = 'C43'
        elif account == 'Discover':
            cell = 'C41'  
        elif account == 'HE_HSA':
            cell = 'P14'
        elif account == 'Liquid Assets':
            cell = 'P4'
        elif account == 'Vanguard401k':
            cell = 'P6'
        elif account == 'VanguardPension':
            cell = 'P8'        
    elif month == 4:
        if account == 'Amex':
            cell = 'K42'
        elif account == 'Barclays':
            cell = 'K45'        
        elif account == 'BoA':
            cell = 'K40'
        elif account == 'BoA-joint':
            cell = 'K52'
        elif account == 'Bonds':
            cell = 'F26'
        elif account == 'Chase':
            cell = 'K43'
        elif account == 'Discover':
            cell = 'K41'  
        elif account == 'HE_HSA':
            cell = 'B36'
        elif account == 'Liquid Assets':
            cell = 'B26'
        elif account == 'Vanguard401k':
            cell = 'B28'
        elif account == 'VanguardPension':
            cell = 'B30'        
    elif month == 5:
        if account == 'Amex':
            cell = 'S42'
        elif account == 'Barclays':
            cell = 'S45'        
        elif account == 'BoA':
            cell = 'S40'
        elif account == 'BoA-joint':
            cell = 'S52'
        elif account == 'Bonds':
            cell = 'M26'
        elif account == 'Chase':
            cell = 'S43'
        elif account == 'Discover':
            cell = 'S41'  
        elif account == 'HE_HSA':
            cell = 'I36'
        elif account == 'Liquid Assets':
            cell = 'I26'
        elif account == 'Vanguard401k':
            cell = 'I28'
        elif account == 'VanguardPension':
            cell = 'I30'        
    elif month == 6:
        if account == 'Amex':
            cell = 'C77'
        elif account == 'Barclays':
            cell = 'C80'        
        elif account == 'BoA':
            cell = 'C75'
        elif account == 'BoA-joint':
            cell = 'C88'
        elif account == 'Bonds':
            cell = 'T26'
        elif account == 'Chase':
            cell = 'C78'
        elif account == 'Discover':
            cell = 'C76'  
        elif account == 'HE_HSA':
            cell = 'P36'
        elif account == 'Liquid Assets':
            cell = 'P26'
        elif account == 'Vanguard401k':
            cell = 'P28'
        elif account == 'VanguardPension':
            cell = 'P30'        
    elif month == 7:
        if account == 'Amex':
            cell = 'K77'
        elif account == 'Barclays':
            cell = 'K80'        
        elif account == 'BoA':
            cell = 'K75'
        elif account == 'BoA-joint':
            cell = 'K88'
        elif account == 'Bonds':
            cell = 'F48'
        elif account == 'Chase':
            cell = 'K78'
        elif account == 'Discover':
            cell = 'K76'  
        elif account == 'HE_HSA':
            cell = 'B58'
        elif account == 'Liquid Assets':
            cell = 'B48'
        elif account == 'Vanguard401k':
            cell = 'B50'
        elif account == 'VanguardPension':
            cell = 'B52'        
    elif month == 8:
        if account == 'Amex':
            cell = 'S77'
        elif account == 'Barclays':
            cell = 'S80'        
        elif account == 'BoA':
            cell = 'S75'
        elif account == 'BoA-joint':
            cell = 'S88'
        elif account == 'Bonds':
            cell = 'M48'
        elif account == 'Chase':
            cell = 'S78'
        elif account == 'Discover':
            cell = 'S76'  
        elif account == 'HE_HSA':
            cell = 'I58'
        elif account == 'Liquid Assets':
            cell = 'I48'
        elif account == 'Vanguard401k':
            cell = 'I50'
        elif account == 'VanguardPension':
            cell = 'I52'        
    elif month == 9:
        if account == 'Amex':
            cell = 'C112'
        elif account == 'Barclays':
            cell = 'C115'        
        elif account == 'BoA':
            cell = 'C110'
        elif account == 'BoA-joint':
            cell = 'C124'
        elif account == 'Bonds':
            cell = 'T48'
        elif account == 'Chase':
            cell = 'C113'
        elif account == 'Discover':
            cell = 'C111'  
        elif account == 'HE_HSA':
            cell = 'P58'
        elif account == 'Liquid Assets':
            cell = 'P48'
        elif account == 'Vanguard401k':
            cell = 'P50'
        elif account == 'VanguardPension':
            cell = 'P52'        
    elif month == 10:
        if account == 'Amex':
            cell = 'K112'
        elif account == 'Barclays':
            cell = 'K115'        
        elif account == 'BoA':
            cell = 'K110'
        elif account == 'BoA-joint':
            cell = 'K124'
        elif account == 'Bonds':
            cell = 'F70'
        elif account == 'Chase':
            cell = 'K113'
        elif account == 'Discover':
            cell = 'K111'  
        elif account == 'HE_HSA':
            cell = 'B80'
        elif account == 'Liquid Assets':
            cell = 'B70'
        elif account == 'Vanguard401k':
            cell = 'B72'
        elif account == 'VanguardPension':
            cell = 'B74'        
    elif month == 11:
        if account == 'Amex':
            cell = 'S112'
        elif account == 'Barclays':
            cell = 'S115'        
        elif account == 'BoA':
            cell = 'S110'
        elif account == 'BoA-joint':
            cell = 'S124'
        elif account == 'Bonds':
            cell = 'M70'
        elif account == 'Chase':
            cell = 'S113'
        elif account == 'Discover':
            cell = 'S111'  
        elif account == 'HE_HSA':
            cell = 'I80'
        elif account == 'Liquid Assets':
            cell = 'I70'
        elif account == 'Vanguard401k':
            cell = 'I72'
        elif account == 'VanguardPension':
            cell = 'I74'        
    elif month == 12:
        if account == 'Amex':
            cell = 'C7'
        elif account == 'Barclays':
            cell = 'C10'        
        elif account == 'BoA':
            cell = 'C5'
        elif account == 'BoA-joint':
            cell = 'C16'
        elif account == 'Bonds':
            cell = 'T70'
        elif account == 'Chase':
            cell = 'C8'
        elif account == 'Discover':
            cell = 'C6'  
        elif account == 'HE_HSA':
            cell = 'P80'
        elif account == 'Liquid Assets':
            cell = 'P70'
        elif account == 'Vanguard401k':
            cell = 'P72'
        elif account == 'VanguardPension':
            cell = 'P74'    
    return cell

def getStartAndEndOfPreviousMonth(today, month, year):
    if month == 1:
        startdate = today.replace(month=12, day=1, year=year - 1)
        enddate = today.replace(month=12, day=31, year=year - 1)
    elif month == 3:
        startdate = today.replace(month=2, day=1)
        enddate = today.replace(month=2, day=28)
    elif month == 5:
        startdate = today.replace(month=4, day=1)
        enddate = today.replace(month=4, day=30)
    elif month == 7:
        startdate = today.replace(month=6, day=1)
        enddate = today.replace(month=6, day=30)
    elif month == 10:
        startdate = today.replace(month=9, day=1)
        enddate = today.replace(month=9, day=30)
    elif month == 12:
        startdate = today.replace(month=11, day=1)
        enddate = today.replace(month=11, day=30)
    else:
        startdate = today.replace(month=month - 1, day=1)
        enddate = today.replace(month=month - 1, day=31)
    return [startdate, enddate]
