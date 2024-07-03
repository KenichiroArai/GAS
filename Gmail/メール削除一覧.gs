/**
 * スプレッドシートからメールアドレスの一覧を取得する。
 * @return メールアドレスの一覧
 */
function getEmailListOfSs() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getDataRange();
  var data = sheet.getDataRange().getValues();

  var emailList = [];
  for (var i = 0; i < data.length; i++) {
    var email = data[i][0];
    if (!email) {
      continue;
    }
    emailList.push(email);
  }
  return emailList;
}

/**
 * Gmailからメールアドレスに一致するデータを削除する。
 * @param emailAddress メールアドレス
 */
function deleteEmailOfGmail(emailAddress) {
  var gmailApp = GmailApp;    // GmailAppオブジェクトを取得
  // 検索条件を作成
  var threads = gmailApp.search("from:" + emailAddress);
  console.log(threads.length);

  for (var i = 0; i < threads.length; i++) {
    // 検索結果のスレッドを削除
    try {
      threads[i].moveToTrash();
    } catch (e) {
      console.log(threads.length - (i + 1));
      console.error(e.message);
      Utilities.sleep(10000);
    }
  }

  return threads.length;
}

/**
 * スプレッドシートからメールアドレスに一致するデータを削除する。
 * @param emailAddress メールアドレス
 */
function deleteEmailOfSs(emailAddress) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    for (var j = 0; j < data[i].length; j++) {
      if (data[i][j] != emailAddress) {
        continue;
      }
      // 値が一致した場合、その行を削除します
      sheet.deleteRow(i + 1);
    }
  }
}

/* メールアドレスの一覧に該当するメールアドレスをGmailから削除する。 */
var emailList = [];
while ((emailList = getEmailListOfSs()).length > 0) {
  const emailAddress = emailList[0];
  console.log(emailAddress + "：開始");
  var length = 0;

  do {
    length = deleteEmailOfGmail(emailAddress);

    Utilities.sleep(10000);
  } while(length > 0);

  deleteEmailOfSs(emailAddress);
  console.log(emailAddress + "：終了");
}
