const name = "Your name here"
const email = "youremailhere@gmail.com"


function putMonthlySubscription() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sh = ss.getSheetByName("Savings")
  var sub = ss.getSheetByName("Subscription")

  var date = Utilities.formatDate(new Date(), "GMT+7", "MM/dd/yyyy")

  var subs = sub.getRange("A2:D4").getValues()
  var range = sh.getRange("G3:G").getValues()

  var filtered = range.filter(String).length + 3
  var values = []
  var index = filtered - 2
  subs.forEach(sub => {
        values.push([index, sub[0] + " subscription", date, sub[1], "Out", 0 , sub[2], sub[3]])
        index = index + 1
  })

  for (i = 0; i < values.length; i++) {
    sh.getRange("G" + (filtered + i) + ":N" + (filtered+i)).setValues([values[i]])
  }
}

function putMonthlyRemuneration() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName("Savings");
  var sub = ss.getSheetByName("Subscription")

  var date = Utilities.formatDate(new Date(), "GMT+7", "MM/dd/yyyy")

  var subs = sub.getRange("F2:H2").getValues()
  var range = sh.getRange("G3:G").getValues();

  var filtered = range.filter(String).length + 3;
  var values = []
  var index = filtered - 2

  subs.forEach(sub => {
        values.push([index, sub[0], date, "Out", sub[1], 0, sub[2], "Transfer"])
        index = index + 1
  })


  for (i = 0; i < values.length; i++) {
    sh.getRange("G" + (filtered + i) + ":N" + (filtered+i)).setValues([values[i]])
  }
}

function generateMonthlyReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sh = ss.getSheetByName("Monthly")
  var date = Utilities.formatDate(new Date(), "GMT+7", "MM/dd/yyyy")

  var range = sh.getRange("A1:A").getValues()
  var filtered = range.filter(String).length + 2

  sh.getRange("A" + (filtered-1) + ":I" + (filtered-1)).setValues([[date, "Debit", "Credit", "Total", "Food", "Entertainment", "Transportation", "Hobby", "Book"]]);

  var values = [
    "Savings", "Cash and E-money", "Investment"
  ]

  var formulas_debit = [
  "=SUMIFS(Savings!$M$3:$M$40000, Savings!$K$3:$K$40000, \"SA-*\", Savings!$I$3:$I$40000, \"<=\"& EOMONTH($A$" + (filtered-1) + ",0), Savings!$I$3:$I$40000, \">=\"& EOMONTH($A$" + (filtered -1) + ",-1)+1)",
  "=SUMIFS('Cash and E-money'!$M$3:$M$40000, 'Cash and E-money'!$K$3:$K$40000, \"CA-*\", 'Cash and E-money'!$I$3:$I$40000, \"<=\"& EOMONTH($A$"+ (filtered-1) + ",0), 'Cash and E-money'!$I$3:$I$40000, \">=\"& EOMONTH($A$"+ (filtered-1) +",-1)+1) + SUMIFS(Savings!$M$3:$M$40000, Savings!$K$3:$K$40000, \"CA-*\", Savings!$I$3:$I$40000, \"<=\"& EOMONTH($A$"+ (filtered-1) + ",0), Savings!$I$3:$I$40000, \">=\"& EOMONTH($A$"+ (filtered-1) + ",-1)+1)",
  "=SUMIFS(Investment!$N$3:$N$40000, Investment!$L$3:$L$40000, \"IN-*\", Investment!$J$3:$J$40000, \"<=\"& EOMONTH($A$"+ (filtered-1) +",0), Investment!$J$3:$J$40000, \">=\"& EOMONTH($A$" + (filtered-1) +",-1)+1) + SUMIFS(Savings!$M$3:$M$40000, Savings!$K$3:$K$40000, \"IN-*\", Savings!$I$3:$I$40000, \"<=\"& EOMONTH($A$" + (filtered-1) +",0), Savings!$I$3:$I$40000, \">=\"& EOMONTH($A$" + (filtered-1) +",-1)+1)"
  ]

  var formulas_credit = [
  "=(SUMIFS(Savings!$M$3:$M$40000, Savings!$J$3:$J$40000, \"SA-*\", Savings!$I$3:$I$40000, \"<=\"& EOMONTH($A$" + (filtered-1) +",0), Savings!$I$3:$I$40000, \">=\"& EOMONTH($A$" + (filtered-1) +",-1)+1) + SUMIFS(Savings!$L$3:$L$40000, Savings!$J$3:$J$40000, \"SA-*\", Savings!$I$3:$I$40000, \"<=\"& EOMONTH($A$" + (filtered-1) + ",0), Savings!$I$3:$I$40000, \">=\"& EOMONTH($A$" + (filtered-1) + ",-1)+1))",
  "=SUMIFS('Cash and E-money'!$M$3:$M$40000, 'Cash and E-money'!$J$3:$J$40000, \"CA-*\", 'Cash and E-money'!$I$3:$I$40000, \"<=\"& EOMONTH($A$"  + (filtered-1) +",0), 'Cash and E-money'!$I$3:$I$40000, \">=\"& EOMONTH($A$"  + (filtered-1) + ",-1)+1)",
  "=SUMIFS(Investment!$N$3:$N$40000, Investment!$K$3:$K$40000, \"IN-*\", Investment!$J$3:$J$40000, \"<=\"& EOMONTH($A$"  + (filtered-1) +",0), Investment!$J$3:$J$40000, \">=\"& EOMONTH($A$"  + (filtered-1) +",-1)+1)"
  ]

  var formula_food = "=SUMIFS('Cash and E-money'!$M$3:$M$40000, 'Cash and E-money'!$J$3:$J$40000, \"CA-*\", 'Cash and E-money'!$I$3:$I$40000, \"<=\"& EOMONTH($A$"  + (filtered-1) +",0), 'Cash and E-money'!$I$3:$I$40000, \">=\"& EOMONTH($A$"  + (filtered-1) + ",-1)+1, 'Cash and E-money'!$N$3:$N$40000, \"Food\") + (SUMIFS(Savings!$M$3:$M$40000, Savings!$J$3:$J$40000, \"SA-*\", Savings!$I$3:$I$40000, \"<=\"& EOMONTH($A$" + (filtered-1) +",0), Savings!$I$3:$I$40000, \">=\"& EOMONTH($A$" + (filtered-1) +",-1)+1, 'Savings'!$N$3:$N$40000, \"Food\") + SUMIFS(Savings!$L$3:$L$40000, Savings!$J$3:$J$40000, \"SA-*\", Savings!$I$3:$I$40000, \"<=\"& EOMONTH($A$" + (filtered-1) + ",0), Savings!$I$3:$I$40000, \">=\"& EOMONTH($A$" + (filtered-1) + ",-1)+1, 'Savings'!$N$3:$N$40000, \"Food\"))"
  var formula_entertainment = "=SUMIFS('Cash and E-money'!$M$3:$M$40000, 'Cash and E-money'!$J$3:$J$40000, \"CA-*\", 'Cash and E-money'!$I$3:$I$40000, \"<=\"& EOMONTH($A$"  + (filtered-1) +",0), 'Cash and E-money'!$I$3:$I$40000, \">=\"& EOMONTH($A$"  + (filtered-1) + ",-1)+1, 'Cash and E-money'!$N$3:$N$40000, \"Entertainment\") + (SUMIFS(Savings!$M$3:$M$40000, Savings!$J$3:$J$40000, \"SA-*\", Savings!$I$3:$I$40000, \"<=\"& EOMONTH($A$" + (filtered-1) +",0), Savings!$I$3:$I$40000, \">=\"& EOMONTH($A$" + (filtered-1) +",-1)+1, 'Savings'!$N$3:$N$40000, \"Entertainment\") + SUMIFS(Savings!$L$3:$L$40000, Savings!$J$3:$J$40000, \"SA-*\", Savings!$I$3:$I$40000, \"<=\"& EOMONTH($A$" + (filtered-1) + ",0), Savings!$I$3:$I$40000, \">=\"& EOMONTH($A$" + (filtered-1) + ",-1)+1, 'Savings'!$N$3:$N$40000, \"Entertainment\"))"
  var formula_transportation = "=SUMIFS('Cash and E-money'!$M$3:$M$40000, 'Cash and E-money'!$J$3:$J$40000, \"CA-*\", 'Cash and E-money'!$I$3:$I$40000, \"<=\"& EOMONTH($A$"  + (filtered-1) +",0), 'Cash and E-money'!$I$3:$I$40000, \">=\"& EOMONTH($A$"  + (filtered-1) + ",-1)+1, 'Cash and E-money'!$N$3:$N$40000, \"Transportation\") + (SUMIFS(Savings!$M$3:$M$40000, Savings!$J$3:$J$40000, \"SA-*\", Savings!$I$3:$I$40000, \"<=\"& EOMONTH($A$" + (filtered-1) +",0), Savings!$I$3:$I$40000, \">=\"& EOMONTH($A$" + (filtered-1) +",-1)+1, 'Savings'!$N$3:$N$40000, \"Transportation\") + SUMIFS(Savings!$L$3:$L$40000, Savings!$J$3:$J$40000, \"SA-*\", Savings!$I$3:$I$40000, \"<=\"& EOMONTH($A$" + (filtered-1) + ",0), Savings!$I$3:$I$40000, \">=\"& EOMONTH($A$" + (filtered-1) + ",-1)+1, 'Savings'!$N$3:$N$40000, \"Transportation\"))"
  var formula_hobby = "=SUMIFS('Cash and E-money'!$M$3:$M$40000, 'Cash and E-money'!$J$3:$J$40000, \"CA-*\", 'Cash and E-money'!$I$3:$I$40000, \"<=\"& EOMONTH($A$"  + (filtered-1) +",0), 'Cash and E-money'!$I$3:$I$40000, \">=\"& EOMONTH($A$"  + (filtered-1) + ",-1)+1, 'Cash and E-money'!$N$3:$N$40000, \"Hobby\") + (SUMIFS(Savings!$M$3:$M$40000, Savings!$J$3:$J$40000, \"SA-*\", Savings!$I$3:$I$40000, \"<=\"& EOMONTH($A$" + (filtered-1) +",0), Savings!$I$3:$I$40000, \">=\"& EOMONTH($A$" + (filtered-1) +",-1)+1, 'Savings'!$N$3:$N$40000, \"Hobby\") + SUMIFS(Savings!$L$3:$L$40000, Savings!$J$3:$J$40000, \"SA-*\", Savings!$I$3:$I$40000, \"<=\"& EOMONTH($A$" + (filtered-1) + ",0), Savings!$I$3:$I$40000, \">=\"& EOMONTH($A$" + (filtered-1) + ",-1)+1, 'Savings'!$N$3:$N$40000, \"Hobby\"))"
  var formula_book = "=SUMIFS('Cash and E-money'!$M$3:$M$40000, 'Cash and E-money'!$J$3:$J$40000, \"CA-*\", 'Cash and E-money'!$I$3:$I$40000, \"<=\"& EOMONTH($A$"  + (filtered-1) +",0), 'Cash and E-money'!$I$3:$I$40000, \">=\"& EOMONTH($A$"  + (filtered-1) + ",-1)+1, 'Cash and E-money'!$N$3:$N$40000, \"Book\") + (SUMIFS(Savings!$M$3:$M$40000, Savings!$J$3:$J$40000, \"SA-*\", Savings!$I$3:$I$40000, \"<=\"& EOMONTH($A$" + (filtered-1) +",0), Savings!$I$3:$I$40000, \">=\"& EOMONTH($A$" + (filtered-1) +",-1)+1, 'Savings'!$N$3:$N$40000, \"Book\") + SUMIFS(Savings!$L$3:$L$40000, Savings!$J$3:$J$40000, \"SA-*\", Savings!$I$3:$I$40000, \"<=\"& EOMONTH($A$" + (filtered-1) + ",0), Savings!$I$3:$I$40000, \">=\"& EOMONTH($A$" + (filtered-1) + ",-1)+1, 'Savings'!$N$3:$N$40000, \"Book\"))"

  for (i = 0; i < values.length; i++) {
    sh.getRange("A" + (filtered + i) + ":A" + (filtered+i)).setValue(values[i]);
    sh.getRange("B" + (filtered + i) + ":B" + (filtered+i)).setFormula(formulas_debit[i]);
    sh.getRange("C" + (filtered + i) + ":C" + (filtered+i)).setFormula(formulas_credit[i]);
    sh.getRange("D" + (filtered + i) + ":D" + (filtered+i)).setFormula("=B" + (filtered + i) + "-C" + (filtered + i));

    if (i == 0) {
      sh.getRange("E" + (filtered + i) + ":E" + (filtered+i)).setFormula(formula_food);
      sh.getRange("F" + (filtered + i) + ":F" + (filtered+i)).setFormula(formula_entertainment);
      sh.getRange("G" + (filtered + i) + ":G" + (filtered+i)).setFormula(formula_transportation);
      sh.getRange("H" + (filtered + i) + ":H" + (filtered+i)).setFormula(formula_hobby);
      sh.getRange("I" + (filtered + i) + ":I" + (filtered+i)).setFormula(formula_book);
    }
  }

  sh.getRange("A" + (filtered + values.length) + ":A" + (filtered + values.length)).setValue("Total");
  sh.getRange("B" + (filtered + values.length) + ":B" + (filtered + values.length)).setFormula("=SUM(B" + filtered + ":B" + (filtered + values.length - 1) + ")");
  sh.getRange("C" + (filtered + values.length) + ":C" + (filtered + values.length)).setFormula("=SUM(C" + filtered + ":C" + (filtered + values.length - 1) + ")");
  sh.getRange("D" + (filtered + values.length) + ":D" + (filtered + values.length)).setFormula("=SUM(D" + filtered + ":D" + (filtered + values.length - 1) + ")");
}

function sendMonthlyReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName("Monthly");
  var date = new Date()

  const months = [
    'January',
    'February',
    'March',
    'April',
    'May',
    'June',
    'July',
    'August',
    'September',
    'October',
    'November',
    'December'
  ]

  const quotes = [
    "Opportunities come infrequently. When it rains gold, put out the bucket, not the thimble.",
    "Price is what you pay. Value is what you get.",
    "Widespread fear is your friend as an investor because it serves up bargain purchases.",
    "Whether we're talking about socks or stocks, I like buying quality merchandise when it is marked down.",
    "We simply attempt to be fearful when others are greedy and to be greedy only when others are fearful.",
    "It's far better to buy a wonderful company at a fair price than a fair company at a wonderful price.",
    "For the investor, a too-high purchase price for the stock of an excellent company can undo the effects of a subsequent decade of favorable business developments.",
    "Someone's sitting in the shade today because someone planted a tree a long time ago",
    "If you aren't willing to own a stock for ten years, don't even think about owning it for ten minutes.",
    "When we own portions of outstanding businesses with outstanding managements, our favorite holding period is forever.",
    "An investor should act as though he had a lifetime decision card with just twenty punches on it.",
    "Since I know of no way to reliably predict market movements, I recommend that you purchase Berkshire shares only if you expect to hold them for at least five years. Those who seek short-term profits should look elsewhere.",
    "Buy a stock the way you would buy a house. Understand and like it such that you'd be content to own it in the absence of any market.",
    "All there is to investing is picking good stocks at good times and staying with them as long as they remain good companies.",
    "Do not take yearly results too seriously. Instead, focus on four or five-year averages.",
    "I never attempt to make money on the stock market. I buy on the assumption that they could close the market the next day and not reopen it for five years.",
    "It is a terrible mistake for investors with long-term horizons -- among them pension funds, college endowments, and savings-minded individuals -- to measure their investment 'risk' by their portfolio's ratio of bonds to stocks,",
    "Never invest in a business you cannot understand.",
    "Risk comes from not knowing what you're doing.",
    "If you don't feel comfortable making a rough estimate of the asset's future earnings, just forget it and move on.",
    "Buy companies with strong histories of profitability and with a dominant business franchise.",
    "We want products where people feel like kissing you instead of slapping you."
  ]

  const monthName = months[date.getMonth()]

  var formatter = new Intl.NumberFormat('en-US',
                                        {style: 'currency',
                                         currency: 'IDR'});


  var range = sh.getRange("A1:A").getValues();
  var startRow = range.filter(String).length - 3;
  var numRows = 4
  var dataRange = sh.getRange(startRow, 2, numRows, 3)
  var data = dataRange.getValues();

  var categoryRange = sh.getRange(startRow, 5, 1, 5)
  var categories = categoryRange.getValues();


  for (var i in data) {
    var row = data[i]
    for (var j in row) {
      data[i][j] = formatter.format(data[i][j])
    }
  }

  var category = categories[0]
  for (var i in category) {
      category[i] = formatter.format(category[i])
   }

  var report = {
    name : name,
    month: monthName + " " + date.getFullYear(),
    savings: data[0],
    cash: data[1],
    investment: data[2],
    total: data[3],
    by_category: category,
    quote: quotes[Math.floor(Math.random() * quotes.length)]
  }

  var templ = HtmlService.createTemplateFromFile('monthly_report');
  templ.report = report;
  var message = templ.evaluate().getContent();

  var emailAddress = email;
  var subject = 'Monthly Financial Report - ' + monthName + " " + date.getFullYear();

  MailApp.sendEmail({
    to: emailAddress,
    subject: subject,
    htmlBody:message
  });
}

function getGrabFoodReceipt(date) {
    const query = "from:(no-reply@grab.com) subject:(Your Grab E-Receipt) after:" + date + " Grabfood";

    let threads = GmailApp.search(query);

    let label = GmailApp.getUserLabelByName("checked");
    if (!label) {label = GmailApp.createLabel("checked")}

    let messages = [];

    threads.forEach(thread => {
                    thread.getMessages().forEach(message => {
                                                 messages.push(message.getPlainBody());
                                                 })
        label.addToThread(thread);
        GmailApp.markThreadRead(thread)
    });
    return messages;
}


function parseGrabfoodEmail(content) {
  let totalRegex = /(?<=RP )(.*)(?= TANGGAL )/
  let paymentAccountRegex = /(?<=Detail Pembayaran:\n)(.*)(?=\nDetail)/
  let transactionTypeRegex = /(?<=Jenis Kendaraan: \n)(.*)(?=\nDiterbitkan oleh)/
  let gotPointRegex = /(?<=Poin yang didapat:\n\+)(.*)(?= poin)/

  let total = content.match(totalRegex)
  let account = content.match(paymentAccountRegex)
  let transactionType = content.match(transactionTypeRegex)
  let gotPoint = content.match(gotPointRegex)
  var point = '0'

  if (gotPoint) {
    point = gotPoint[0]
  }

  let transaction = {
    "total" :total[0],
    "account_type" : account[0].trim(),
    "transaction_type": transactionType[0].trim(),
    "point": point
  }

  return transaction
}

function putDailyGrabfoodReceiptEmail() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName("Cash And E-money");
  var date = Utilities.formatDate(new Date(), "GMT+7", "MM/dd/yyyy")

  var range = sh.getRange("G3:G").getValues();
  var filtered = range.filter(String).length + 3;

  let messages = getGrabFoodReceipt(date)

  var values = []
  var index = filtered - 2
  messages.forEach(message => {
        let account = "CA-4";
        let transaction = parseGrabfoodEmail(message);
        if (transaction.account_type == "OVO Points") {
            account = "CA-5"
        }
        values.push([index, transaction.transaction_type, date, account, "Out", 0 , transaction.total, "Food"])
        index = index + 1;
        if (transaction.point !== '0') {
          values.push([index, transaction.transaction_type + ' Point', date, "Out", "CA-5", 0 , transaction.point, "Others"])
          index = index + 1
        }
    });

  for (i = 0; i < values.length; i++) {
    sh.getRange("G" + (filtered + i) + ":N" + (filtered+i)).setValues([values[i]]);
  }
}