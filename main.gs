/**
 * Extract users with administrative privileges and users with inappropriate password settings.
 * @param none
 * @return none
 */
function getGoogleAccountInformation(){
  /****** Start setting constants ******/
  const settingSheetName = '設定用シート';
  const passwordSheetName = 'パスワードの長さに関するコンプライアンスチェック';
  const adminSheetName = '管理者権限チェック';
  const allSheetName = '全て';
  const outputSheetsName = [passwordSheetName, adminSheetName, allSheetName];
  const adminStatusColName = '管理者のステータス';
  const passwordLengthCheckColName = 'パスワードの長さに関するコンプライアンス';
  const passwordSafetyCheckColName = 'パスワードの安全度';
  const targetHeaders = ['ユーザー', 'ユーザー アカウントのステータス', adminStatusColName, '2 段階認証プロセスの登録', '2 段階認証プロセスの適用',passwordLengthCheckColName , passwordSafetyCheckColName];
  const ss_url = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(settingSheetName).getRange('B1').getValue();
  /****** End setting constants ******/
  // If there is no output sheet, create an output sheet.
  outputSheetsName.forEach(x => {
    if (!SpreadsheetApp.getActiveSpreadsheet().getSheetByName(x)){
      SpreadsheetApp.getActiveSpreadsheet().insertSheet();
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().setName(x);
      setTargetColumnsWidth(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(x));
    }
  });
  const allSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(allSheetName); 
  const ss = SpreadsheetApp.openByUrl(ss_url);
  const sheet = ss.getSheets()[0];
  const rawUserLogs = sheet.getDataRange().getValues();
  // Edit the headers.
  const temp1UserLogs = rawUserLogs.map((x, idx) => idx == 0? x.map(y => y.replace(/ \[.*$/u, '')): x);
  // Transpose the array and extract only the columns I need.
  const transposeFunction = temp1UserLogs => temp1UserLogs[0].map((x, idx) => temp1UserLogs.map(y => y[idx]));
  const temp2UserLogs = transposeFunction(temp1UserLogs);
  const temp3UserLogs = temp2UserLogs.filter(x => targetHeaders.indexOf(x[0]) > -1);
  const targetValues = transposeFunction(temp3UserLogs);
  // Output to sheet.
  // All values.
  setValuesToSheet(allSheet, targetValues);
  // Password check.
  const passwordLengthCheckCol = targetValues[0].indexOf(passwordLengthCheckColName);
  const passwordSafetyCheckCol = targetValues[0].indexOf(passwordSafetyCheckColName);
  const passwordCheck = targetValues.filter((x, idx) => x[passwordLengthCheckCol] == '非準拠' || x[passwordLengthCheckCol] == '不明' || x[passwordSafetyCheckCol] == '弱' || x[passwordSafetyCheckCol] == '不明' || idx == 0);
  setValuesToSheet(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(passwordSheetName), passwordCheck);
  // Admin.
  const adminStatusCol = targetValues[0].indexOf(adminStatusColName);
  const adminStatus = targetValues.filter((x, idx) => x[adminStatusCol] == '特権管理者' || x[adminStatusCol] == '管理者' || idx == 0);
  setValuesToSheet(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(adminSheetName), adminStatus);
  // Hide sheets that are not to be included in the PDF output.
  const hiddenStatus = hideSheetNonTargetPrinting(SpreadsheetApp.getActiveSpreadsheet(), outputSheetsName);
  // Output PDF.
  const filename = 'ISF27-4 Google Workspace権限確認（別表）_' + Utilities.formatDate(new Date(), 'JST','yyyyMMdd');
  convertSpreadsheetToPdf(SpreadsheetApp.getActiveSpreadsheet(),
                          null,
                          false,
                          2,
                          filename,
                          DriveApp.getRootFolder()
  );
  // Restores the visible/hidden status of the sheet.
  hiddenStatus.forEach(x => x[1]? SpreadsheetApp.getActiveSpreadsheet().getSheetByName(x[0]).hideSheet(): SpreadsheetApp.getActiveSpreadsheet().getSheetByName(x[0]).showSheet());
}
/**
 * Set array values to cells.
 * @param {sheet} Output sheet.
 * @param {Array.<string>} Values to output.
 * @return none.
 */
function setValuesToSheet(sheet, targetValues){
  sheet.clear();
  sheet.getRange(1, 1).setValue('出力日：' + Utilities.formatDate(new Date(), 'JST','yyyy/M/d'));
  sheet.getRange(2, 1, targetValues.length, targetValues[0].length).setValues(targetValues);
  setTargetColumnsWidth(sheet);
  sheet.setFrozenRows(2);
}
/**
 * Set columns width.
 * @parem {sheet} Target sheet.
 * @return none.
 */
function setTargetColumnsWidth(sheet){
  /* Adjust the width of the email address column. 
     The function "autoResizeColumn" cannot adjust the width of columns with Japanese characters.*/
  sheet.autoResizeColumn(1);
  sheet.setColumnWidth(2, 210);
  sheet.setColumnWidth(3, 128);
  sheet.setColumnWidth(4, 165);
  sheet.setColumnWidth(5, 165);
  sheet.setColumnWidth(6, 275);
  sheet.setColumnWidth(7, 128);
}
function onOpen() {
  SpreadsheetApp.getActiveSpreadsheet().addMenu('権限チェック', [{name:'権限チェック', functionName:'getGoogleAccountInformation'}]);
}