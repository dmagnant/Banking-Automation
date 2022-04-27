import csv
from datetime import datetime, timedelta
import gspread
from Functions import openGnuCashBook, setDirectory, getStartAndEndOfPreviousMonth, chromeDriverAsUser, showMessage

def compileGnuTransactions(account, mybook, directory, date_range):
    import_csv = directory + r"\Projects\Coding\Python\BankingAutomation\Resources\import.csv"
    open(import_csv, 'w', newline='').truncate()
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
    gnu_account = matchAccount()
    
    total = 0
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
                total += abs(float(amount))
                csv.writer(open(import_csv, 'a', newline='')).writerow(row)
    return total

def getDateRange(lastDay):
    # Gather last 3 days worth of transactions
    current_date = lastDay.date()
    date_range = current_date.isoformat()
    num = 1
    month = lastDay.month
    while month == lastDay.month:
        day_before = (current_date - timedelta(days=num))
        num += 1
        month = day_before.month
        if month == lastDay.month:
            date_range = date_range + day_before.isoformat()
    return date_range

def updateSpreadsheet(directory, account, month, value, accounts='p'):
    json_creds = directory + r"\Projects\Coding\Python\BankingAutomation\Resources\creds.json"
    sheetTitle = 'Asset Allocation' if accounts == 'p' else 'Home'
    sheet = gspread.service_account(filename=json_creds).open(sheetTitle)
    worksheetTitle = 'Goals' if accounts == 'p' else 'Finances'
    worksheet = sheet.worksheet(worksheetTitle)
    cell = getCell(account, month, accounts)
    worksheet.update(cell, value)

def getCell(account, month, accounts='p'):
    row_start = 48 if accounts == 'p' else 25
    row = str(row_start + (month - 1))
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
    date_range = getDateRange(lastmonth[1])
    month = lastmonth[1].month

    directory = setDirectory()
    # Set Gnucash Book
    mybook = openGnuCashBook(directory, 'Finance', True, True) if accounts == 'p' else openGnuCashBook(directory, 'Home', False, False)

    if accounts == 'p': 
        JointExpenses = compileGnuTransactions('Joint Expenses', mybook, directory, date_range)
        CCRewards = compileGnuTransactions('CC Rewards', mybook, directory, date_range)
        Dividends = compileGnuTransactions('Dividends', mybook, directory, date_range)
        Interest = compileGnuTransactions('Interest', mybook, directory, date_range)
        MarketResearch = compileGnuTransactions('Market Research', mybook, directory, date_range)
        
        updateSpreadsheet(directory, 'Joint Expenses', month, JointExpenses)
        updateSpreadsheet(directory, 'CC Rewards', month, CCRewards)
        updateSpreadsheet(directory, 'Dividends', month, Dividends)
        updateSpreadsheet(directory, 'Interest', month, Interest)
        updateSpreadsheet(directory, 'Market Research', month, MarketResearch)
    else:
        Groceries = compileGnuTransactions('Groceries', mybook, directory, date_range)
        HomeDepot = compileGnuTransactions('Home Depot', mybook, directory, date_range)
        HomeExpenses = compileGnuTransactions('Home Expenses', mybook, directory, date_range)
        HomeFurnishings = compileGnuTransactions('Home Furnishings', mybook, directory, date_range)
        Pet = compileGnuTransactions('Pet', mybook, directory, date_range)
        Travel = compileGnuTransactions('Travel', mybook, directory, date_range)
        Utilities = compileGnuTransactions('Utilities', mybook, directory, date_range)
        Dan = compileGnuTransactions('Dan', mybook, directory, date_range)
        Tessa = compileGnuTransactions('Tessa', mybook, directory, date_range)

        updateSpreadsheet(directory, 'Groceries', month, Groceries, accounts)
        updateSpreadsheet(directory, 'Home Depot', month, HomeDepot, accounts)
        updateSpreadsheet(directory, 'Home Expenses', month, HomeExpenses, accounts)
        updateSpreadsheet(directory, 'Home Furnishings', month, HomeFurnishings, accounts)
        updateSpreadsheet(directory, 'Pet', month, Pet, accounts)
        updateSpreadsheet(directory, 'Travel', month, Travel, accounts)
        updateSpreadsheet(directory, 'Utilities', month, Utilities, accounts)
        updateSpreadsheet(directory, 'Dan', month, Dan, accounts)
        updateSpreadsheet(directory, 'Tessa', month, Tessa, accounts)

    Amazon = compileGnuTransactions('Amazon', mybook, directory, date_range)
    BarsRestaurants = compileGnuTransactions('Bars & Restaurants', mybook, directory, date_range)
    Entertainment = compileGnuTransactions('Entertainment', mybook, directory, date_range)
    Other = compileGnuTransactions('Other', mybook, directory, date_range)

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

runUpdateMonthlyGoals('j')
