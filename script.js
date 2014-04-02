// config options
var SPREADSHEET_ID = 'ADD_YOUR_SPREADSHEET_ID_HERE';

function main() {
  // get accout name or CID
  var accountName = AdWordsApp.currentAccount().getName();
  if(accountName === null) {
    accountName = AdWordsApp.currentAccount().getCustomerId();
  }
  
  // open or creeat a new sheet for the account
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(accountName);
  if(sheet == null) {
    var sheet = ss.insertSheet(accountName);
  }
  
  // get all enabled ad groups using Ad Group Performance Report
  var awql = "SELECT Id, Name, CampaignName FROM ADGROUP_PERFORMANCE_REPORT WHERE CampaignStatus = ACTIVE AND AdGroupStatus = ENABLED DURING LAST_MONTH";
  var adGroupsToCheck = {};
  var adGroupIds = [];
  var adGroupRows = AdWordsApp.report(awql).rows();
  while (adGroupRows.hasNext()) {
    var adGroup = adGroupRows.next();
    adGroupsToCheck[adGroup['Id']] = {
      id: adGroup['Id'],
      name: adGroup['Name'],
      campaignName: adGroup['CampaignName'],
      adCount: 0
    }
    adGroupIds.push(adGroup['Id']);
  }
  
  // set ad count for each ad group using Ad Performance Report
  var adGroupIdList = "[" + adGroupIds.join(", ") + "]";
  awql = "SELECT AdGroupId FROM AD_PERFORMANCE_REPORT WHERE Status = ENABLED AND AdGroupId IN " + adGroupIdList + " DURING LAST_MONTH";
  var adRows = AdWordsApp.report(awql).rows();
  while (adRows.hasNext()) {
    var ad = adRows.next();
    adGroupsToCheck[ad['AdGroupId']].adCount += 1;
  }
  
  // go through each ad group in the object and create a simple vector to be copied in the sheet
   var adGroupList = [];
  for (var adGroupId in adGroupsToCheck) {
    var adGroup = adGroupsToCheck[adGroupId];
    if (adGroup.adCount < 3 || adGroup.adCount > 5) {
      adGroupList.push([adGroup.campaignName, adGroup.name, adGroup.adCount]);
    }
  }
  var adGroupLength = adGroupList.length;
  
  // get last month date
  var date = new Date();
  var lastMonth = new Date(date.getFullYear(), date.getMonth() - 1);
  
  // set header labels
  sheet.getRange('A1:D1').setValues([['Month', 'Campaign Name', 'AdGroup', 'Ad Count']]).setBackground('#93c47d').setFontWeight('bold');
  
  // make some space for new data
  sheet.insertRows(2, adGroupLength + 1);
  
  // do some beautify and formatting of the sheet
  sheet.getRange(2, 1, adGroupLength, 1).merge().setValue(lastMonth).setBackground('#fff2cc').setHorizontalAlignment('center').setVerticalAlignment('middle').setFontWeight('bold').setNumberFormat('MMMM yyyy');
  sheet.setFrozenRows(1);
  
  // write the ad groups data
  sheet.getRange(2, 2, adGroupLength, 3).setValues(adGroupList);
  
  // auto resize all columns
  sheet.autoResizeColumn(1).autoResizeColumn(2).autoResizeColumn(3).autoResizeColumn(4);
}