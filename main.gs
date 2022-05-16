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
  const adminStatusColName = '管理者のステータス';
  const passwordLengthCheckColName = 'パスワードの長さに関するコンプライアンス';
  const passwordSafetyCheckColName = 'パスワードの安全度';
  const targetHeaders = ['ユーザー', 'ユーザー アカウントのステータス', adminStatusColName, '2 段階認証プロセスの登録', '2 段階認証プロセスの適用',passwordLengthCheckColName , passwordSafetyCheckColName];
  const ss_url = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(settingSheetName).getRange('B1').getValue();
  /****** End setting constants ******/
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
  setValuesToSheet_(allSheet, targetValues);
  // Password check.
  const passwordLengthCheckCol = targetValues[0].indexOf(passwordLengthCheckColName);
  const passwordSafetyCheckCol = targetValues[0].indexOf(passwordSafetyCheckColName);
  const passwordCheck = targetValues.filter((x, idx) => x[passwordLengthCheckCol] == '非準拠' || x[passwordLengthCheckCol] == '不明' || x[passwordSafetyCheckCol] == '弱' || x[passwordSafetyCheckCol] == '不明' || idx == 0);
  setValuesToSheet_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(passwordSheetName), passwordCheck);
  // Admin.
  const adminStatusCol = targetValues[0].indexOf(adminStatusColName);
  const adminStatus = targetValues.filter((x, idx) => x[adminStatusCol] == '特権管理者' || x[adminStatusCol] == '管理者' || idx == 0);
  setValuesToSheet_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(adminSheetName), adminStatus);
  // Output PDF.
  [
    ['ISF27-4 管理台帳レビュー記録 資料 Googleアカウント ' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd'), [adminSheetName, allSheetName]], 
    ['ISF27-5 パスワードポリシー遵守確認 GoogleWorkspace', [passwordSheetName]]
  ].forEach(x => outputPdf_(x[0], x[1]));
}
/**
 * Output PDF.
 * @param {String} Output file name.
 * @param {Array.<string>} An array of sheet names to output.
 * @return none.
 */
function outputPdf_(filename, outputSheetNames){
  createOutputSheets_(outputSheetNames);
  // Hide sheets that are not to be included in the PDF output.
  const hiddenStatus = hideSheetNonTargetPrinting(SpreadsheetApp.getActiveSpreadsheet(), outputSheetNames);
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
 * If there is no output sheet, create an output sheet.
 * @param {Array.<string>} An array of sheet names to output.
 * @return none.
 */
function createOutputSheets_(outputSheetsName){
  outputSheetsName.forEach(x => {
    if (!SpreadsheetApp.getActiveSpreadsheet().getSheetByName(x)){
      SpreadsheetApp.getActiveSpreadsheet().insertSheet();
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().setName(x);
      setTargetColumnsWidth_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(x));
    }
  });
}
/**
 * Set array values to cells.
 * @param {sheet} Output sheet.
 * @param {Array.<string>} Values to output.
 * @return none.
 */
function setValuesToSheet_(sheet, targetValues){
  sheet.clear();
  sheet.getRange(1, 1).setValue('出力日：' + Utilities.formatDate(new Date(), 'JST','yyyy/M/d'));
  sheet.getRange(2, 1, targetValues.length, targetValues[0].length).setValues(targetValues);
  setTargetColumnsWidth_(sheet);
  sheet.setFrozenRows(2);
}
/**
 * Set columns width.
 * @parem {sheet} Target sheet.
 * @return none.
 */
function setTargetColumnsWidth_(sheet){
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
function onOpen(){
  SpreadsheetApp.getActiveSpreadsheet().addMenu('権限チェック', [{name:'権限チェック', functionName:'getGoogleAccountInformation'}]);
}