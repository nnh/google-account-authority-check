/**
* Create a PDF.
* @param {spreadsheet} Spreadsheet to output.
* @param {string} Name of the sheet to output.
* @param {boolean} Output orientation. true:vertical false:Horizontal
* @param {number} Output Scale. 1= 100%(default), 2= Fit to width, 3= Fit to height,  4= Fit to page.
* @param {string} Name of the PDF to output.
* @param {folder} Output folder of GoogleDrive.
* @return none
*/
function convertSpreadsheetToPdf(ss, sheetName, portrait, scale, pdfName, outputFolder){
  const urlBase = ss.getUrl().replace(/edit.*$/,'');
  let strId = '&id=' +ss.getId();
  if (sheetName != null){
    const sheetId = ss.getSheetByName(sheetName).getSheetId();
    strId = '&gid=' + sheetId;
  }
  const urlExport = 'export?exportFormat=pdf&format=pdf'
      + strId
      + '&size=A4'
      + '&portrait=' + portrait 
      + '&fitw=true'
      + '&scale=' + scale
      + '&sheetnames=true'
      + '&printtitle=false'
      + '&pagenum=CENTER'
      + '&gridlines=false'  // hide gridlines
      + '&fzr=true';        // repeat row headers (frozen rows) on each page
  const options = {
    headers: {
      'Authorization': 'Bearer ' +  ScriptApp.getOAuthToken(),
    }
  }
  const response = UrlFetchApp.fetch(urlBase + urlExport, options);
  const blob = response.getBlob().setName(pdfName + '.pdf');
  outputFolder.createFile(blob);
}
/**
 * Hide sheets that are not to be printed. Returns the original show/hide status.
 * @param {spreadsheet} target spreadsheet.
 * @param {Array.<string>} Names of sheet to print.
 * @return {Array.<string, boolean>} Sheet name, visible: False, hidden: True.
 */
function hideSheetNonTargetPrinting(ss, outputSheetsName){
  const sheets = ss.getSheets();
  // Get the sheet show/hide status, True if hidden.
  const sheetVisibleStatus = sheets.map(x => [x.getName(), x.isSheetHidden()]);
  sheets.forEach(x => {
    if (outputSheetsName.indexOf(x.getName()) == -1){
      x.hideSheet();
    }
  });
  return sheetVisibleStatus;
}