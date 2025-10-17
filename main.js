class GetGoogleAccountInformationAll {
  constructor() {
    this.adminStatusColName = "管理者のステータス";
    this.passwordLengthCheckColName =
      "パスワードの長さに関するコンプライアンス";
    this.passwordSafetyCheckColName = "パスワードの安全度";
    this.targetHeaders = [
      "ユーザー",
      "ユーザー アカウントのステータス",
      this.adminStatusColName,
      "2 段階認証プロセスの登録",
      "2 段階認証プロセスの適用",
      this.passwordLengthCheckColName,
      this.passwordSafetyCheckColName,
    ];
    this.outputSheetName = "全て";
    this.sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      this.outputSheetName
    );
    this.outputFilename =
      "ISF27-4 管理台帳レビュー記録 資料 Googleアカウント " +
      Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyyMMdd");
  }
  getInputData_() {
    const settingSheetName = "設定用シート";
    const ss_url = SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName(settingSheetName)
      .getRange("B1")
      .getValue();
    const ss = SpreadsheetApp.openByUrl(ss_url);
    const rawUserLogs = ss.getSheets()[0].getDataRange().getValues();
    // Edit the headers.
    const userLogEditHeader = rawUserLogs.map((x, idx) =>
      idx == 0 ? x.map((y) => y.replace(/ \[.*$/u, "")) : x
    );
    const tempUserLogOnlyOutputCol = this.transposeFunction_(userLogEditHeader);
    const userLogOnlyOutputCol = this.getOnlyOutputCol_(
      tempUserLogOnlyOutputCol,
      this.targetHeaders
    );
    return userLogOnlyOutputCol;
  }
  transposeFunction_(target) {
    return target[0].map((_, idx) => target.map((y) => y[idx]));
  }
  getOnlyOutputCol_(inputData, targetColNames) {
    const onlyTargetCol = inputData.filter(
      (x) => targetColNames.indexOf(x[0]) > -1
    );
    return this.transposeFunction_(onlyTargetCol);
  }
  editInputData_() {
    return this.getInputData_();
  }
  /**
   * Set array values to cells.
   * @param none.
   * @return none.
   */
  setValuesToSheet_() {
    const targetValues = this.editInputData_();
    this.sheet.clear();
    this.sheet
      .getRange(1, 1)
      .setValue(
        "出力日：" + Utilities.formatDate(new Date(), "JST", "yyyy/M/d")
      );
    this.sheet
      .getRange(2, 1, targetValues.length, targetValues[0].length)
      .setValues(targetValues);
    this.setTargetColumnsWidth_();
    this.sheet.setFrozenRows(2);
  }
  /**
   * Set columns width.
   * @parem none.
   * @return none.
   */
  setTargetColumnsWidth_() {
    /* Adjust the width of the email address column. 
       The function "autoResizeColumn" cannot adjust the width of columns with Japanese characters.*/
    this.sheet.autoResizeColumn(1);
    this.sheet.setColumnWidth(2, 210);
    this.sheet.setColumnWidth(3, 128);
    this.sheet.setColumnWidth(4, 165);
    this.sheet.setColumnWidth(5, 165);
    this.sheet.setColumnWidth(6, 275);
    this.sheet.setColumnWidth(7, 128);
  }
  /**
   * Output PDF.
   * @param none.
   * @return none.
   */
  outputPdf_() {
    this.setValuesToSheet_();
    this.createOutputSheets_(this.outputSheetName);
    // Hide sheets that are not to be included in the PDF output.
    const hiddenStatus = hideSheetNonTargetPrinting(
      SpreadsheetApp.getActiveSpreadsheet(),
      [this.outputSheetName]
    );
    convertSpreadsheetToPdf(
      SpreadsheetApp.getActiveSpreadsheet(),
      null,
      false,
      2,
      this.outputFilename,
      DriveApp.getRootFolder()
    );
    // Restores the visible/hidden status of the sheet.
    hiddenStatus.forEach((x) =>
      x[1]
        ? SpreadsheetApp.getActiveSpreadsheet().getSheetByName(x[0]).hideSheet()
        : SpreadsheetApp.getActiveSpreadsheet().getSheetByName(x[0]).showSheet()
    );
  }
  /**
   * If there is no output sheet, create an output sheet.
   * @param none.
   * @return none.
   */
  createOutputSheets_() {
    if (
      !SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
        this.outputSheetName
      )
    ) {
      SpreadsheetApp.getActiveSpreadsheet().insertSheet();
      SpreadsheetApp.getActiveSpreadsheet()
        .getActiveSheet()
        .setName(this.outputSheetName);
      this.setTargetColumnsWidth_(
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
          this.outputSheetName
        )
      );
    }
  }
}
class GetGoogleAccountInformationPasswordPolicy extends GetGoogleAccountInformationAll {
  constructor() {
    super();
    this.outputSheetName = "パスワードの長さに関するコンプライアンスチェック";
    this.sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      this.outputSheetName
    );
    this.outputFilename =
      "ISF27-5 パスワードポリシー遵守確認 GoogleWorkspace";
  }
  editInputData_() {
    const targetValues = this.getInputData_();
    const passwordLengthCheckCol = targetValues[0].indexOf(
      this.passwordLengthCheckColName
    );
    const passwordSafetyCheckCol = targetValues[0].indexOf(
      this.passwordSafetyCheckColName
    );
    const passwordCheck = targetValues.filter(
      (x, idx) =>
        x[passwordLengthCheckCol] == "非準拠" ||
        x[passwordLengthCheckCol] == "不明" ||
        x[passwordSafetyCheckCol] == "弱" ||
        x[passwordSafetyCheckCol] == "不明" ||
        idx == 0
    );
    return passwordCheck;
  }
}
class GetGoogleAccountInformationAdmin extends GetGoogleAccountInformationAll {
  constructor() {
    super();
    this.outputSheetName = "管理者権限チェック";
    this.sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      this.outputSheetName
    );
  }
  editInputData_() {
    const targetValues = this.getInputData_();
    const adminStatusCol = targetValues[0].indexOf(this.adminStatusColName);
    const adminStatus = targetValues.filter(
      (x, idx) =>
        x[adminStatusCol] == "特権管理者" ||
        x[adminStatusCol] == "管理者" ||
        idx == 0
    );
    return adminStatus;
  }
}
function getGoogleAccountInformationPasswordPolicy() {
  const GetGoogleInfoAll = new GetGoogleAccountInformationAll();
  // Output to sheet.
  // All values.
  GetGoogleInfoAll.outputPdf_();
  const GetGoogleInfoPasswordPolicy =
    new GetGoogleAccountInformationPasswordPolicy();
  // Password policy.
  GetGoogleInfoPasswordPolicy.outputPdf_();
}
function getGoogleAccountInformationAdmin() {
  const GetGoogleInfoAdmin = new GetGoogleAccountInformationAdmin();
  GetGoogleInfoAdmin.setValuesToSheet_();
}
function onOpen() {
  SpreadsheetApp.getActiveSpreadsheet().addMenu(
    "パスワードポリシー遵守確認",
    [
      {
        name: "PDF出力",
        functionName: "getGoogleAccountInformationPasswordPolicy",
      },
    ]
  );
  SpreadsheetApp.getActiveSpreadsheet().addMenu("アクセス権管理棚卸", [
    { name: "シート出力", functionName: "getGoogleAccountInformationAdmin" },
  ]);
}
