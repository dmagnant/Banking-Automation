import csv
from datetime import datetime, timedelta
import gspread
from Functions import openGnuCashBook, setDirectory, getStartAndEndOfPreviousMonth, chromeDriverAsUser, showMessage

def compileGnuTransactions(account, mybook, directory, dateRange):
    importCSV = directory + r"\Projects\Coding\Python\BankingAutomation\Resources\import.csv"
    open(importCSV, 'w', newline='').truncate()
    def matchAccount():
        match account:
            case 'Amazon':
                return 'Expenses:Amazon'
            case 'Bars & Restaurants':
                return 'Expenses:Bars & Restaurants'
            case 'CC Rewards':
                return 'Income:Credit Card Rewards'
            case 'Dan':
                return "Dan's Contributions"
            case 'Dividends':
                return 'Income:Investments:Dividends'
            case 'Entertainment':
                return 'Expenses:Entertainment'
            case 'Groceries':
                return 'Expenses:Groceries'
            case 'Home Depot':
                return 'Expenses:Home Depot'
            case 'Home Expenses':
                return 'Expenses:Home Expenses'
            case 'Home Furnishings':
                return 'Expenses:Home Furnishings'
            case 'Interest':
                return 'Income:Investments:Interest'
            case 'Joint Expenses':
                return 'Expenses:Joint Expenses'
            case 'Market Research':
                return 'Income:Market Research'
            case 'Other':
                return 'Expenses:Other'
            case 'Pet':
                return 'Expenses:Pet'
            case 'Tessa':
                return "Tessa's Contributions"
            case 'Travel':
                return 'Expenses:Travel'
            case 'Utilities':
                return 'Expenses:Utilities'
    gnuAccount = matchAccount()
    
    total = 0
    # retrieve transactions from GnuCash
    transactions = [tr for tr in mybook.transactions
                    if str(tr.post_date.strftime('%Y-%m-%d')) in dateRange
                    for spl in tr.splits
                    if spl.account.fullname == gnuAccount]
    for tr in transactions:
        date = str(tr.post_date.strftime('%Y-%m-%d'))
        description = str(tr.description)
        for spl in tr.splits:
            amount = format(spl.value, ".2f")
            if spl.account.fullname == gnuAccount:
                row = date, description, str(amount)
                total += abs(float(amount))
                csv.writer(open(importCSV, 'a', newline='')).writerow(row)
    return total

def getDateRange(lastDay):
    # Gather last 3 days worth of transactions
    currentDate = lastDay.date()
    dateRange = currentDate.isoformat()
    num = 1
    month = lastDay.month
    while month == lastDay.month:
        dayBefore = (currentDate - timedelta(days=num))
        num += 1
        month = dayBefore.month
        if month == lastDay.month:
            dateRange = dateRange + dayBefore.isoformat()
    return dateRange

def updateSpreadsheet(directory, account, month, value, accounts='p'):
    jsonCreds = directory + r"\Projects\Coding\Python\BankingAutomation\Resources\creds.json"
    sheetTitle = 'Asset Allocation' if accounts == 'p' else 'Home'
    sheet = gspread.service_account(filename=jsonCreds).open(sheetTitle)
    worksheetTitle = 'Goals' if accounts == 'p' else 'Finances'
    worksheet = sheet.worksheet(worksheetTitle)
    cell = getCell(account, month, accounts)
    worksheet.update(cell, value)

def getCell(account, month, accounts='p'):
    rowStart = 48 if accounts == 'p' else 25
    row = str(rowStart + (month - 1))
    match account:
        case 'Amazon':
            return 'C' + row if accounts == 'p' else 'B' + row
        case 'Bars & Restaurants':
            return 'D' + row if accounts == 'p' else 'C' + row
        case 'CC Rewards':
            return 'L' + row
        case 'Dan':
            return 'Q' + row
        case 'Dividends':
            return 'M' + row
        case 'Entertainment':
            return 'E' + row if accounts == 'p' else 'D' + row
        case 'Groceries':
            return 'E' + row
        case 'Home Depot':
            return 'F' + row
        case 'Home Expenses':
            return 'G' + row
        case 'Home Furnishings':
            return 'H' + row
        case 'Interest':
            return 'N' + row
        case 'Joint Expenses':
            return 'F' + row
        case 'Market Research':
            return 'O' + row
        case 'Other':
            return 'G' + row if accounts == 'p' else 'J' + row
        case 'Pet':
            return 'K' + row
        case 'Tessa':
            return 'R' + row         
        case 'Travel':
            return 'L' + row
        case 'Utilities':
            return 'M' + row

def runUpdateMonthlyGoals(accounts):
    # accounts = p for personal or j for joint
    # get current date
    today = datetime.today()
    year = today.year
    month = today.month
    lastmonth = getStartAndEndOfPreviousMonth(today, month, year)
    dateRange = getDateRange(lastmonth[1])
    month = lastmonth[1].month

    directory = setDirectory()
    # Set Gnucash Book
    mybook = openGnuCashBook(directory, 'Finance', True, True) if accounts == 'p' else openGnuCashBook(directory, 'Home', False, False)

    if accounts == 'p': 
        JointExpenses = compileGnuTransactions('Joint Expenses', mybook, directory, dateRange)
        CCRewards = compileGnuTransactions('CC Rewards', mybook, directory, dateRange)
        Dividends = compileGnuTransactions('Dividends', mybook, directory, dateRange)
        Interest = compileGnuTransactions('Interest', mybook, directory, dateRange)
        MarketResearch = compileGnuTransactions('Market Research', mybook, directory, dateRange)
        
        updateSpreadsheet(directory, 'Joint Expenses', month, JointExpenses)
        updateSpreadsheet(directory, 'CC Rewards', month, CCRewards)
        updateSpreadsheet(directory, 'Dividends', month, Dividends)
        updateSpreadsheet(directory, 'Interest', month, Interest)
        updateSpreadsheet(directory, 'Market Research', month, MarketResearch)
    else:
        Groceries = compileGnuTransactions('Groceries', mybook, directory, dateRange)
        HomeDepot = compileGnuTransactions('Home Depot', mybook, directory, dateRange)
        HomeExpenses = compileGnuTransactions('Home Expenses', mybook, directory, dateRange)
        HomeFurnishings = compileGnuTransactions('Home Furnishings', mybook, directory, dateRange)
        Pet = compileGnuTransactions('Pet', mybook, directory, dateRange)
        Travel = compileGnuTransactions('Travel', mybook, directory, dateRange)
        Utilities = compileGnuTransactions('Utilities', mybook, directory, dateRange)
        Dan = compileGnuTransactions('Dan', mybook, directory, dateRange)
        Tessa = compileGnuTransactions('Tessa', mybook, directory, dateRange)

        updateSpreadsheet(directory, 'Groceries', month, Groceries, accounts)
        updateSpreadsheet(directory, 'Home Depot', month, HomeDepot, accounts)
        updateSpreadsheet(directory, 'Home Expenses', month, HomeExpenses, accounts)
        updateSpreadsheet(directory, 'Home Furnishings', month, HomeFurnishings, accounts)
        updateSpreadsheet(directory, 'Pet', month, Pet, accounts)
        updateSpreadsheet(directory, 'Travel', month, Travel, accounts)
        updateSpreadsheet(directory, 'Utilities', month, Utilities, accounts)
        updateSpreadsheet(directory, 'Dan', month, Dan, accounts)
        updateSpreadsheet(directory, 'Tessa', month, Tessa, accounts)

    Amazon = compileGnuTransactions('Amazon', mybook, directory, dateRange)
    BarsRestaurants = compileGnuTransactions('Bars & Restaurants', mybook, directory, dateRange)
    Entertainment = compileGnuTransactions('Entertainment', mybook, directory, dateRange)
    Other = compileGnuTransactions('Other', mybook, directory, dateRange)

    updateSpreadsheet(directory, 'Amazon', month, Amazon, accounts)
    updateSpreadsheet(directory, 'Bars & Restaurants', month, BarsRestaurants, accounts)
    updateSpreadsheet(directory, 'Entertainment', month, Entertainment, accounts)
    updateSpreadsheet(directory, 'Other', month, Other, accounts)

    directory = setDirectory()
    driver = chromeDriverAsUser(directory)
    driver.implicitly_wait(3)
    # Display Asset Allocation spreadsheet
    driver.get('https://docs.google.com/spreadsheets/d/1sWJuxtYI-fJ6bUHBWHZTQwcggd30RcOSTMlqIzd1BBo/edit#gid=1813404638') if accounts == 'p' else driver.get('https://docs.google.com/spreadsheets/d/1oP3U7y8qywvXG9U_zYXgjFfqHrCyPtUDl4zPDftFCdM/edit#gid=1436385671')
    showMessage('Review Spreadsheet', 'Once complete, click OK to close')
    driver.close()

if __name__ == '__main__':
    directory = setDirectory()
    driver = chromeDriverAsUser()
    accounts = 'j'
    runUpdateMonthlyGoals(accounts)
