var formSheet = SpreadsheetApp.openById("[ID REMOVED]");
var email, employeeName, dateOfPurchase;
const today = new Date();
var error = "";
var currentRow = "";
const folderIDs = ["1ToglUeFxYm7RiG3_oR0xJ6_BdD6fgXh2", "1DUCCs8YM_4IvjMO20BA4Q5E03Nv2cUgT", "1_b6klntjRwyD0uXJnRK7PzFwmqjZfYhK", "1nNti_zgvoyzsXtYzyKeEjOk6Nkg5oA8C", "1cJWcU7-CVDgZREWbazs-0thC32afiPdE", "1IHFwb2eUK5U4uQOMX3A0XejdXfjOmzvG", "1d2JeK8_HQAnQWxNbC2RLCDiDQN1lGe6O", "1ZO1FVTXtqLjPZmvv4-12FRsoR9LnDbH_", "1C0VPW8X9qJZ5VTTNMrI93xpnpLrJJhKw", "1NrWN8UXV53nO7EWOSFpCAvw301dIDFza", "1DO1Z2IvWUx7995cYqOEQuWmqcQ0V6ru7"];
if (today.getDate() >= 8 && today.getMonth() != parseInt(PropertiesService.getScriptProperties().getProperty("currentMonth"))) {
  PropertiesService.getScriptProperties().setProperty("currentMonth", today.getMonth());
  PropertiesService.getScriptProperties().setProperty("sheetNumOfMonth", parseInt(PropertiesService.getScriptProperties().getProperty("sheetNumOfMonth")) + 1);
  console.log(`property "sheetNumOfMonth" was changed to ${PropertiesService.getScriptProperties().getProperty("sheetNumOfMonth")}`);  
}
var monthNum = parseInt(PropertiesService.getScriptProperties().getProperty("sheetNumOfMonth"));
var latestEntry = [];
const defaultMsg = "<p>Thank you for submitting your reimbursement form. If the form was submitted correctly, please look for your reimbursement on the payroll that is sent on the 15th of the month. If you had questions that require a response, we will respond via email. Thanks again for your hard work and dedication to Bnos Yisroel.</p>";
const lateMsg = "<p>Thank you for submitting your reimbursement form. Your form has not been processed because it was submitted more than 30 (thirty) days after your purchase. If you would like to discuss this further,  please speak to Mrs. Heyman. Thanks again for your hard work and dedication to Bnos Yisroel.</p>";

function printFolderIDs() {
  var folder = DriveApp.getFolderById("1RLLH3T9XbI8qZlHWaBo7plVk-bqXbwKRk1wQ2_ganPcnxZe4hTD_m363JumUVgmzrcTTj70W")
  var subfolders = folder.getFolders() 
  while(subfolders.hasNext()) {
    var subfolder=subfolders.next();
    console.log(`${subfolder.getName()}: ${subfolder.getId()}`);
  }
}

function manualSend() {
  var rowToSend = 20;
  currentRow = rowToSend;
  var sheetData = formSheet.getSheets()[0].getDataRange().getValues();
  for (var col = 0; col < 13; col++) {
    if (col == 11) {continue;}
      latestEntry.push(sheetData[rowToSend - 1][col]);
    }
  var e = {values: latestEntry};
  main(e);
}

function main(e) {
  // ScriptApp.newTrigger("main").forSpreadsheet(formSheet).onFormSubmit().create();
  latestEntry = e.values;
  for (var i in latestEntry) {
    console.log(`${i}: ${latestEntry[i]}`);
  }  
  employeeName = `${latestEntry[2]} ${latestEntry[1]}`;
  email = latestEntry[3];
  dateOfPurchase = new Date(latestEntry[5]);
  if (dateOfPurchase > today) {
    error = "Error - Date entered was in the future";
    sendEmail(`<p>There was an error with your form. The date you entered for your purchase (${dateOfPurchase.toDateString()}) is in the future. Please fill out the form again: ${getLink()}<br />Thank you!</p>`);
    return;    
  }
  if (isEarlierThan30Days(dateOfPurchase)) {
    error = "Error - Job was more than 30 days ago";
    sendEmail(lateMsg); 
    return;   
  }
  renameAndMovePics();
  sendEmail(defaultMsg);
}

function logInThisMonth() {
  if (currentRow == "") {currentRow = formSheet.getSheets()[0].getLastRow();}
  formSheet.getSheets()[0].getRange("L" + currentRow).setValue(error);
  latestEntry.push(latestEntry[11]);
  latestEntry[11] = error;
  const lastRow = formSheet.getSheets()[monthNum].getLastRow() + 1;
  formSheet.getSheets()[monthNum].getRange(`A${lastRow}:M${lastRow}`).setValues([latestEntry]);
}

function renameAndMovePics() {
  var links = latestEntry[11].split(", ");
  var file;
  for (var i in links) {
    file = DriveApp.getFileById(getID(links[i])).setName(`${dateOfPurchase.getFormDate()}, ${latestEntry[1]}, ${latestEntry[2]}, ${latestEntry[4]}`);
    file.moveTo(DriveApp.getFolderById(folderIDs[0]));
    file.makeCopy(file.getName(), DriveApp.getFolderById(folderIDs[monthNum]));
  }
}

function sendEmail(html) {
  logInThisMonth();  
  var body = `<p>Regarding form submitted by: ${employeeName}</p>${html}`;
  var labels = formSheet.getSheets()[0].getDataRange().getValues();
    for (var col = 0; col < 13; col++) {
      if (latestEntry[col] != "") {body += `${labels[0][col]}: ${latestEntry[col]}<br />`};
    }
  GmailApp.sendEmail(email, "Bnos Yisroel Reimbursement", "", {
    name: "Bnos Yisroel",  
    htmlBody: body,
    bcc: "[EMAIL REMOVED]"
    });
    console.log(`Sent: ${body}`);
    console.log(`Remaining emails: ${MailApp.getRemainingDailyQuota()}`);
}

function isEarlierThan30Days(dateToCalc) {
  const diffTime = Math.abs(today - dateToCalc);
  return Math.ceil(diffTime / (1000 * 60 * 60 * 24)) > 30;
}

function getID(link) {
  return link.substring(link.indexOf("id=") + 3);
}

function getLink() {
return `https://docs.google.com/forms/d/e/1FAIpQLSfRQiESHP5iQvuDRstTOzeP7pwMkUHN2EyrpL3pHvscfjV1pw/viewform?usp=pp_url&entry.301288006=${latestEntry[1].replaceAll(" ", "+")}&entry.8578947=${latestEntry[2].replaceAll(" ", "+")}&entry.26536188=${latestEntry[3]}&entry.1334154573=${latestEntry[4].replaceAll(" ", "+")}&entry.1943965431=${latestEntry[5].getFormDate()}&entry.618291784=${latestEntry[6].replaceAll(" ", "+")}&entry.1955631681=${latestEntry[7].replaceAll(" ", "+")}&entry.1931654830=${latestEntry[8].replaceAll(" ", "+")}&entry.2142302727=${latestEntry[9].replaceAll(" ", "+")}&entry.176526013=${latestEntry[10].replaceAll(" ", "+")}`
}

function addLeadingZeroIfNone(num) {
  if (num.toString().length == 1) {
    num = "0" + num;
  }
  return num;
}

 String.prototype.getFormDate = function () {
   var str = this.valueOf();
    if (str == "") {return "";}
    else {
      str = str.split("/");
      return `${str[2]}-${addLeadingZeroIfNone(str[1])}-${addLeadingZeroIfNone(str[0])}`;
    }
};

 Date.prototype.getFormDate = function () {
   return `${this.getFullYear()}-${addLeadingZeroIfNone(this.getMonth() + 1)}-${addLeadingZeroIfNone(this.getDate())}`;
 }
