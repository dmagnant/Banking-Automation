# Banking Automation

PROBLEM: 

I found the benefit of using financial aggregator sites was having all of my account balances in one place, but the drawback was reliquinshing my usernames/passwords. I also use spreadsheets to budget/forecast which helps me feel financially secure but I found that manually gathering/entering the data necessary for these was time-consuming and error-prone.

SOLUTION: 

First, I utilized GnuCash, an open-source accounting software, as a replacement for a financial aggregator site. Second, to populate data into both GnuCash and my spreadsheets, I utilized Python and Selenium to login to and download transactions/statements from each of my financial websites and upload them accordingly.

HOW-TO-USE:

Python files were created for each site that, when executed, populate the necessary data into GnuCash and the spreadsheets. Checking and Savings accounts are designed to be checked daily, while other accounts are designed to capture the monthly statement.

 | | NOTABLE FEATURES | | 
 
 - Cross-validation of transactions to avoid duplicates
 
For daily monitoring of checking/savings accounts, I found that I could not simply download one day's worth of transactions each day because in some cases transactions posted overnight and had the prior day's posting date, meaning they would not be captured when downloading transactions for the next day. Therefore, the application downloads multiple days worth of transactions from the checking/savings websites, compares those against the transactions already uploaded to GnuCash, and only uploads new transactions.

 - Pop-up message of any unrecognized transactions

Transactions whose descriptions/categories are not hard-coded within the application will be placed in the "Other" expense category, while also appearing in a pop-up message displaying the transaction details. This not only offers a chance to manually re-categorize, but if the transaction still isn't recognized, this may serve as notice of a fraudulent transaction.

- Automatic redemption of credit card rewards

The application will check the balance of cashback rewards and automatically redeem the points if the minimum is met.
