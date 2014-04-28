var SPREADSHEET_ID = 'ENTER_YOUR_SPREADSHEET_ID';

function main() {
  // sheet labels
  var labels = [['Campaign Name', 'AdGroup Name', 'Headline', 'Headline Count', 'DKI', 'Description 1', 'Desc 1 Count', 'Description 2', 'Desc 2 Count', 'Desc 2 !']];
  
  // open spreadsheet
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // get all adwords account
  var accountSelector = MccApp.accounts()
    .forDateRange("THIS_MONTH")
    .orderBy("Clicks DESC");

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
    }
    
    // clear old data contents and formatting
    sheet.getRange('A2:J').clear();
    
    // vars to store data and background colors
    var adData = new Array(), backgrounds = new Array();
    
    // get all text ads from an account
    var adsIterator = AdWordsApp.ads()
      .forDateRange('THIS_MONTH')
      .withCondition('Type = TEXT_AD')
      .withCondition('CampaignStatus = ACTIVE')
      .withCondition('AdGroupStatus = ENABLED')
      .withCondition('Status = ENABLED')
      .orderBy('CampaignName ASC')
      .get();
    
    //get each ad
    while (adsIterator.hasNext()) {
      var ad = adsIterator.next();
      
      // get headline text and character count
      var headline = ad.getHeadline();
      var headlineSize = (headline.indexOf('{param') !== -1) ? headline.length - 9 : headline.length;
      if(headline.toLowerCase().indexOf('{keyword') !== -1) {
        headlineSize = headline.length - 10;
      }
      // set headline highlight color based on character count
      var headlineSizeColor = (headlineSize <= 21) ? '#f4cccc' : '';
      
      // check headline dynamic keyword insertion
      var headlineDKI = (headline.toLowerCase().indexOf('{keyword') !== -1 && headline.indexOf('{KeyWord') !== -1) ? '' : 'NOT OK';
      if(headline.toLowerCase().indexOf('{keyword') === -1) {
        var headlineDKI = '';
      }
      var headDKIColor = (headlineDKI === 'NOT OK') ? '#f4cccc' : '';
      
      
      // get description 1 text and character count
      var description1 = ad.getDescription1();
      var description1Size = (description1.indexOf('{param') !== -1) ? description1.length - 9 : description1.length;
      
      // set description 1 highlight color based on character count
      var desc1SizeColor = (description1Size <= 31) ? '#f4cccc' : '';
      
      
      // get description 2 text and character count
      var description2 = ad.getDescription2();
      var description2Size = (description2.indexOf('{param') !== -1) ? description2.length - 9 : description2.length;
      
      // set description 2 highlight color based on character count
      var desc2SizeColor = (description2Size <= 31) ? '#f4cccc' : '';
      
      // description 2 exclamation mark
      var desc2Excl = (description2.indexOf('!') === -1) ? 'NOT OK' : '';
      var desc2ExclColor = (desc2Excl === 'NOT OK') ? '#f4cccc' : '';
      
      // get only ads with problems
      if(headlineSize <= 21 || description1Size <= 31 || description2Size <= 31 || desc2Excl === 'NOT OK' || headlineDKI === 'NOT OK') {
        adData.push([ad.getCampaign().getName(), ad.getAdGroup().getName(), headline, headlineSize, headlineDKI, description1, description1Size, description2, description2Size, desc2Excl]);
        backgrounds.push(['', '', '', headlineSizeColor, headDKIColor, '', desc1SizeColor, '', desc2SizeColor, desc2ExclColor]);
      }
    }
    
    // save ads on account sheet
    sheet.getRange('A1:J1').setValues(labels).setBackground('#93c47d').setFontWeight('bold');
    sheet.getRange(2, 1, adData.length, 10).setValues(adData).setBackgrounds(backgrounds);
    sheet.setFrozenRows(1);
  }
}