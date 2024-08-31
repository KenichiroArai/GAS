/**
 * スケジュールをインポートする。
 */
function importSchedule() {

  /* フォルダの指定。全て同じフォルダIDに指定可能 */
  // インポート対象
  const IMPORT_TARGET_FOLDER_ID = "<インポート対象フォルダID>";
  // インポート完了
  const IMPORT_COMPLETED_FOLDER_ID = "<インポート完了フォルダID>";
  // 中間ファイル生成
  const INTERMEDIATE_FILE_GENERATION_FOLDER_ID = "<中間ファイル生成フォルダID>";

  /* カレンダーIDの定義 */
  // カレンダーID
  const CALENDAR_ID = "<カレンダーID>";

  /* ドキュメントを作成する */
  let convertedFileIds = createDocuments(IMPORT_TARGET_FOLDER_ID, INTERMEDIATE_FILE_GENERATION_FOLDER_ID);
  if (convertedFileIds.length <= 0) {
    console.info("インポート対象に該当ファイルがありません。");
    return;
  }

  /* テキストを出力する */
  for (let convertedFileId of convertedFileIds) {
    let text = getText(convertedFileId);
    const fileName = DriveApp.getFileById(convertedFileId).getName() + ".csv";

    // ファイル削除
    deleteFileByName(fileName);

    const fileId = createCalendarImportFile(IMPORT_TARGET_FOLDER_ID, fileName, text, INTERMEDIATE_FILE_GENERATION_FOLDER_ID);
    importCSVtoCalendar(fileId, CALENDAR_ID);
  }

  /* インポート対象のファイルをインポート完了に移動する。 */
  let tagetFileIds = getTagetFileIds(IMPORT_TARGET_FOLDER_ID);
  for (let targetFileId of tagetFileIds) {
    // ファイルの移動
    moveFileToFolder(targetFileId, IMPORT_COMPLETED_FOLDER_ID);
  }

  /**
   * フォルダIDに該当する対象ファイルIDの一覧を取得する。
   * @param {string} folderId フォルダID
   * @return {string[]} イメージファイルID
   */
  function getTagetFileIds(folderId) {
    var result = [];

    let folder = DriveApp.getFolderById(folderId);
    
    let files = folder.getFiles();
    while(files.hasNext()) {
      let file = files.next();

      var mimeType = file.getMimeType();

      // ファイルタイプが画像またはPDFではないか
      if (!(mimeType.startsWith('image/') || mimeType === 'application/pdf')) {
        // 画像またはPDFではない
        continue;
      }

      result.push(file.getId());
    }

    return result;
  }

  /**
   * ファイルIDからファイルを削除する。
   * @param {string} fileId ファイルID
   * @param {string} excludedFileId 対象外ファイルID
   */
  function deleteFileById(fileId, excludedFileId) {
    // ファイルIDからファイルを取得
    var file = DriveApp.getFileById(fileId);
    
    // ファイル名を取得
    var fileName = file.getName();
    
    // ファイル名に該当するファイルを検索
    var files = DriveApp.getFilesByName(fileName);
    
    // 該当するファイルを削除
    while (files.hasNext()) {
      var fileToDelete = files.next();

      // 該当するファイルが対象外か
      if (fileToDelete.getId() == excludedFileId) {
        // 対象外の場合

        continue;
      }

      fileToDelete.setTrashed(true); // ファイルをゴミ箱に移動
    }
  }

  /**
   * ファイル名からファイルを削除する。
   * @param {string} fileName ファイル名
   */
  function deleteFileByName(fileName) {
    var files = DriveApp.getFilesByName(fileName);
    while (files.hasNext()) {
      var file = files.next();
      file.setTrashed(true);
    }
  }


  /**
   * 入力ファイルIDからドキュメントを作成する。
   * @param {string} inputFileId 入力ファイルID
   * @param {string} outputFolderId 出力フォルダID
   * @return {string} ドキュメントID
   */
  function createDocument(inputFileId, outputFolderId) {
    var result = ""

    let ocrOption = {
      "ocr": true,
      "ocrLanguage": "ja",
    };
    let ocrResource = {
      mimeType: MimeType.GOOGLE_DOCS,
    };
    let documentFile = Drive.Files.copy(ocrResource, inputFileId, ocrOption);
    
    // コピー先フォルダにファイルを移動
    DriveApp.getFileById(documentFile.id).moveTo(DriveApp.getFolderById(outputFolderId));

    result = documentFile.id;

    return result;
  }

  /**
   * テキストを取得する。
   * @param {string} documentFileId ドキュメントファイルID
   * @return {string} テキスト
   */
  function getText(documentFileId) {
    let result = "";
    let documentFile = DocumentApp.openById(documentFileId);
    result = documentFile.getBody().getText();
    return result;
  }

  /**
   * 入力フォルダIDに該当するファイル一覧の全てのドキュメントを出力フォルダIDのフォルダ内に作成する。
   * @param {string} inputFolderId 入力フォルダID
   * @param {string} outputFolderId 出力フォルダID
   * @return {string[]} ドキュメントIDの一覧
   */
  function createDocuments(inputFolderId, outputFolderId) {
    let result = [];
    let tagetFileIds = getTagetFileIds(inputFolderId);
    for (let targetFileId of tagetFileIds) {

      let convertedFileId = createDocument(targetFileId, outputFolderId);
      result.push(convertedFileId);
      
      // ファイルを削除する
      deleteFileById(convertedFileId, convertedFileId);
    }
    return result;
  }

  /**
   * カレンダーのインポートファイルを作成する。
   * @param {string} folderId フォルダID
   * @param {string} fileName ファイル名
   * @param {string} contents 中身
   * @param {string} outputFolderId 出力フォルダID
   * @return {string} インポートファイルID
   */
  function createCalendarImportFile(folderId, fileName, contents, outputFolderId) {

    let result = "";

    const PATTERN_DATA = /.(\d+\/\d+) \(.\) (\d+:\d+) ?~? ?(.*)/;
    const PATTERN_HEAD_SQUARE = /^■/;
    const PATTERN_SQUARE = /■/;
    const PATTERN_DATA2 = /(.+)■(\d+\/\d+) \(.\) (\d+:\d+) ?~? ?(.*)/;

    // ファイルに書き込む内容
    let writeContents = "";

    /* 内容を行ごとにファイルに書き込む内容に設定する */
    let contentsArrays = contents.split(/\n/);
    for (line of contentsArrays) {

      // 行にデータが無いか
      if (line == null) {
        // 無い場合

        continue;
      }

      // 行をトリムする
      line_wk = line.trim();
      // データがあるか
      if (line_wk == "") {
        // データが無い場合

        continue;
      }

      // ■が先頭にないか
      if (!line_wk.match(PATTERN_HEAD_SQUARE)) {
        // 先頭にない場合

        let matchResult = line_wk.match(PATTERN_DATA);
        // データパーンに一致しないか
        if (!matchResult) {
          // 一致しない場合

          writeContents += line_wk;

          continue;
        }

        writeContents += getWriteContents(matchResult);

        continue;
      }
      
      const matchResult = line_wk.match(PATTERN_DATA);
      if (!matchResult) {
        continue;
      }

      writeContents += getWriteContents(matchResult);
    }

    writeContents = writeContents.substring(1, writeContents.length);

    result = createCsvFile(folderId, fileName, writeContents, outputFolderId);

    return result;
  }

  /**
   * 書き込み中身を返す。
   * @param {string} matchResult マッチ結果
   * @return {string} 書き込み中身
   */
  function getWriteContents(matchResult) {
    let result = "";

    const PATTERN_SQUARE = /■/;
    const PATTERN_DATA2 = /(.+)■(\d+\/\d+) \(.\) (\d+:\d+) ?~? ?(.*)/;

    let matchResultSquare = matchResult[3].match(PATTERN_SQUARE);
    if (matchResultSquare) {
      matchResult2 = matchResult[3].match(PATTERN_DATA2);
      matchResult[3] = matchResult2[1];

      result += "\n";
      result += matchResult[1] + ", " + matchResult[2] + ", " + matchResult[3] + ", " + matchResult[3];

      result += "\n";
      result += matchResult2[2] + ", " + matchResult2[3] + ", " + matchResult2[4] + ", " + matchResult2[4];

      return result;
    }

    result += "\n";
    result += matchResult[1] + ", " + matchResult[2] + ", " + matchResult[3] + ", " + matchResult[3];

    return result;
  }

  /**
   * CSVファイル作成する。
   * @param {string} folderId フォルダID
   * @param {string} fileName ファイル名
   * @param {string} contents ファイルの内容
   * @param {string} outputFolderId 出力フォルダID
   * @return {string} CSVファイルID
   */
  function createCsvFile(folderId, fileName, contents, outputFolderId) {

    let result = "";

    const contentType = 'text/csv';                     // コンテンツタイプ
    const charset = 'UTF-8';                            // 文字コード
    const folder = DriveApp.getFolderById(folderId);    // 出力するフォルダ

    const blob = Utilities.newBlob('', contentType, fileName).setDataFromString(contents, charset);
    
    const file = folder.createFile(blob);
    result = file.getId();
    
    // ファイルを移動する
    const outputFolder = DriveApp.getFolderById(outputFolderId); // 移動先のフォルダ
    file.moveTo(outputFolder);

    return result;
  }

  /**
   * CSVからカレンダーにインポートする。
   * @param {string} csvFileId CSVファイルID
   * @param {string} calendarId カレンダーID
   */
  function importCSVtoCalendar(csvFileId, calendarId) {

    var calendar = CalendarApp.getCalendarById(calendarId);

    var file = DriveApp.getFileById(csvFileId);
    var csvDatas = Utilities.parseCsv(file.getBlob().getDataAsString());

    for (var line of csvDatas) {
      const title = line[2];
      const year = new Date().getFullYear();
      const date = new Date(year + "/" + line[0]);
      const times = line[1].split(":");
      const hours = times[0];
      date.setHours(hours);
      const minutes = times[1];
      date.setMinutes(minutes);
      const startTime = date;
      const endTime = date;
      const description = line[3];

      console.log("【カレンダーインポートデータ】title:%s, startTime:%s, endTime:%s, description:%s", title, startTime, endTime, description);

      calendar.createEvent(title, startTime, endTime, {
        description: description,
      });
    }
  }

  /**
   * ファイルを移動する。
   * @param {string} fileId ファイルID
   * @param {string} folderId フォルダID
   */
  function moveFileToFolder(fileId, folderId) {
    var file = DriveApp.getFileById(fileId);
    var folder = DriveApp.getFolderById(folderId);
    file.moveTo(folder);
  }

}
