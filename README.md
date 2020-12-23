# chipsy-finsy
Chipsy finsy is simple financial recording templates using google sheets and app script

## Features
1. Automate regular monthly billing (subs, etc)
2. Parse email-based receipt (grab, etc)
3. Create monthly report and send it to your email
4. Monitor current investment value (stocks,mf)

## How to use
1. Go to this sheet template link http://bit.ly/chipsy-finsy-sheet
2. Click File and choose "Make A Copy", this will add the template to your own drive
3. Open "Copy of Financial Sheets Template" sheet file on your own Google drive
4. Click "Tools > Script Editor"
5. Change the name and email constant in Code.gs, (if you can't find Code.gs and monthly_report.html file you can get it here)
6. Done, now you can adjust the sheet to reflect your own financial condition
7. Don't forget to authorize the script

## Sheets Guide
This is the guide to adjust the sheet easily
1. Go to "Savings" sheet and adjust cell A3:D10 to reflect your saving account, adjust cell B17:B24 to adjust initial/current value
2. Go to "Cash and E-money" sheet and adjust cell A3:D16 to reflect your emoney account, adjust cell B23:B36 to adjust initial/current value
3. Go to "Investment" sheet and adjust cell A3:D9 to reflect your investment account, adjust cell B16:B22 to adjust initial/current value
4. Go to "Subscription" sheet and adjust cell A2:D4 to reflect your monthly subscription / spending, adjust cell F2:H2 to adjust your monthly income

## Guide to record your transaction
After finishing the setup you only need to record your transaction,
The only sheet that need to be updated are "Savings" , "Cash and E-money", and "Investment", it serves as transaction log that will automatically update the whole debit credit thingy

## Schedule The Automation
1. Click triggers on the sidebar
2. Click Add Trigger:
   1. Choose which function to run: **putDailyGrabfoodTransaction**
   2. Choose which deployment should run: **Head**
   3. Select event source: **Time-driven**
   4. Select type of time based trigger: **Day timer**
   5. Select time of day: **11pm to midnight**

Above is the example trigger configuration to set daily cron job to extract Grabfood receipt and add it to Cash and E-money transaction log

## Recommended Configuration

| Function                    | Description                                                                                                              | Recommended Trigger |
|-----------------------------|--------------------------------------------------------------------------------------------------------------------------|---------------------|
| putMonthlySpending          | Automatically adding all monthly spending in "Subscription" sheet to transaction log in "Savings" sheet every month      | Month Timer         |
| putMonthlyIncome            | Automatically adding all monthly income in "Subscription" sheet to transaction log in "Savings" sheet every month        | Month Timer         |
| generateMonthlyReport       | Automatically create monthly report every month                                                                          | Month Timer         |
| sendMonthlyReport           | Automatically send monthly report to your email                                                                          | Month Timer         |
| putDailyGrabfoodTransaction | Automatically parse grabfood receipt from your email and add it to transaction log in "Cash and E-money" sheet every day | Day Timer           |
