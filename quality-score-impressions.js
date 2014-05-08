// config options
var SPREADSHEET_ID = 'ENTER_YOUR_SPREADSHEET_ID';

function main() {
  // open spreadsheet
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // get last week dates
  var d = new Date; // get current date
  var firstDate = d.getDate() - d.getDay() - 6;
  var lastDate = firstDate + 6;
  var firstDay = Utilities.formatDate(new Date(d.setDate(firstDate)), "GMT", "yyyy-MM-dd");
  var lastDay = Utilities.formatDate(new Date(d.setDate(lastDate)), "GMT", "yyyy-MM-dd");
  
  // get all adwords accounts
  var accountSelector = MccApp.accounts()
    .forDateRange("LAST_WEEK")
    .orderBy("Clicks DESC").withLimit(1);

  // get each account
  var accountIterator = accountSelector.get();
  while (accountIterator.hasNext()) {
    var account = accountIterator.next();
    
    // activate current account
    MccApp.select(account);

    // get account name or cid (if name is not set)
    var accountName = (AdWordsApp.currentAccount().getName() != '') ? AdWordsApp.currentAccount().getName() : AdWordsApp.currentAccount().getCustomerId();;
    
    // open account sheet, or create a new one
    var sheet = ss.getSheetByName(accountName);
    if(sheet === null) {
      sheet = ss.insertSheet(accountName);
      
      // insert labels
      insertLabels(sheet);
    }
    
    // get current sheet start date
    var startDate = sheet.getRange('B1').getValue();
    if(startDate != '') {
      startDate.setDate(startDate.getDate() + 1);
      startDate = Utilities.formatDate(startDate, "GMT", "yyyy-MM-dd");
    }
    
    // insert new column if we have different week
    if(firstDay > startDate) {
      sheet.insertColumnBefore(2);
    }
    
    // insert start and end date
    sheet.getRange('B1:B2').setValues([[firstDay], [lastDay]]).setFontWeight('bold').setNumberFormat("MMMM d").setBackground('#cccccc');
    sheet.getRange('B17:B18').setValues([[firstDay], [lastDay]]).setFontWeight('bold').setNumberFormat("MMMM d").setBackground('#cccccc');
    
    // store impressions for each quality score
    var qsImpressions = new Array();
    
    // get all ads from an account
    var keywordsIterator = AdWordsApp.keywords()
      .forDateRange('LAST_WEEK')
      .withCondition('Status = ACTIVE')
      .withCondition('Impressions > 0')
      .withCondition("CampaignName CONTAINS_IGNORE_CASE 'search'")
      .get();
    
    // initialize all qs impressions to 0
    for(var i=0; i<=10; i++) {
      qsImpressions[i] = 0;
    }
    
    //get each ad
    while (keywordsIterator.hasNext()) {
      var keyword = keywordsIterator.next();
      var stats = keyword.getStatsFor('LAST_WEEK');
      var impressions = stats.getImpressions();
      
      // sum impressions for each quality score
      qsImpressions[keyword.getQualityScore()] += impressions;
    }
    
    // insert impressions into the sheet
    for(var i in qsImpressions) {
      if(i != 0) {
        var row = parseInt(i) + 18;
        sheet.getRange('B' + row).setValue(qsImpressions[i]);
      }
    }
    
    // insert formulas
    insertFormulas(sheet);
  }
}

// create sheet labels and formatting
function insertLabels(sheet) {
  sheet.getRange('A2').setValue('QS - Pondere Impresii');
  for(var i=1; i<=10; i++) {
    var row = i+2;
    sheet.getRange('A' + row).setValue(i);
  }
  sheet.getRange('A13').setValue('% 7 or more');
  sheet.getRange('A2:A12').setBackground('#fce5cd');
  sheet.getRange('A13:13').setBackground('#b7b7b7');
  
  sheet.getRange('A18').setValue('QS - Total Impresii');
  for(var i=1; i<=10; i++) {
    var row = i+18;
    sheet.getRange('A' + row).setValue(i);
  }
  sheet.getRange('A29').setValue('Total Impresii:');
  sheet.getRange('A18:A29').setBackground('#fce5cd');
}

// insert sheet formulas
function insertFormulas(sheet) {
  sheet.getRange('B13').setFormula('=SUM(B9:B12)');
  
  for(var i=1; i<=10; i++) {
    var row = i+2;
    var val = i+18;
    sheet.getRange('B' + row).setFormula('=(B' + val + '/B29)');
  }
  
  sheet.getRange('B29').setFormula('=SUM(B19:B28)');
  
  sheet.getRange('B3:B13').setNumberFormat('0.00%');
}