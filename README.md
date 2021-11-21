# Banking-Automation

PROBLEM: 

I found the benefit of financial aggregator sites was seeing balances for all my accounts in one place, but the drawback was giving out my credentials for each site to another party. Further, I would also use spreadsheets to budget/forecast, and found that gathering/entering the data necessary for these tasks was time-consuming and error-prone.

SOLUTION: 

First, I utilized GnuCash, an open-source accounting software, as a replacement for the financial aggregator site. Second, to populate data into both GnuCash and my spreadsheets, I utilized Python and Selenium to login to and download transactions/statements from each of my financial websites and upload them accordingly.

HOW-TO-USE:

I created Python files for each site that, when executed, populate the necessary data into GnuCash and my spreadsheets. Checking and Savings accounts are designed to be checked daily, other accounts are designed to capture the monthly statement.

 | | NOTABLE FEATURES | | 
 
 - Cross-validation of transactions to avoid duplicates
 
For daily checking of checking/savings accounts, I found that I could not simply download one day's worth of transactions each day because in some cases, transactions posted overnight yet had the prior day's posting date. Therefore, I designed the application to download multiple days worth of transactions from my account's websites, compare those against the transactions already uploaded to GnuCash, and only upload new transactions.

 - Pop-up message of any unrecognized transactions

Transactions whose descriptions/categories are not hard-coded within the application will be placed in the "Other" expense category, while also appearing in a pop-up message to notify the user. This not only gives the user a chance to manually re-categorize if they recognize the transaction, but if they don't recognize the transaction, this serves as notice of potentially fraudulent activity.

