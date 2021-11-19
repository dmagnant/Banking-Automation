from selenium import webdriver
import socket
from pykeepass import PyKeePass
import os
import psutil
import time
import ctypes
import pygetwindow
import pyautogui
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
        elif account == 'Vanguard':
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