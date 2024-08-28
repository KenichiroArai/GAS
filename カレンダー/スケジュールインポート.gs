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
 * 対象ファイルIDからドキュメントを作成する。
 * @param {string} targetFileId 対象ファイルID
 * @return {string} ドキュメントID
 */
function createDocument(targetFileId) {
  var result = ""

  let ocrOption = {
    "ocr": true,
    "ocrLanguage": "ja",
  };
  let ocrResource = {
    mimeType: MimeType.GOOGLE_DOCS,
  };
  let documentFile = Drive.Files.copy(ocrResource, targetFileId, ocrOption);

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
 * フォルダIDに該当するファイル一覧の全てのドキュメントを作成する。
 * @param {string} folderId フォルダID
 * @return {string[]} ドキュメントIDの一覧
 */
function createDocuments(folderId) {
  let result = [];
  let tagetFileIds = getTagetFileIds(folderId);
  for (let targetFileId of tagetFileIds) {
    let convertedFileId = createDocument(targetFileId);
    result.push(convertedFileId);
  }
  return result;
}

/**
 * カレンダーのインポートファイルを作成する。
 * @param {string} folderId フォルダID
 * @param {string} fileName ファイル名
 * @param {string} contents 中身
 * @return {string} インポートファイルID
 */
function createCalendarImportFile(folderId, fileName, contents) {

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

    const PATTERN_DATA = /.(\d+\/\d+) \(.\) (\d+:\d+)~(.*)/;

    const PATTERN_SQUARE = /^■/;
    // ■が先頭にないか
    if (!line_wk.match(PATTERN_SQUARE)) {
      // 先頭にない場合

      const matchResult = line_wk.match(PATTERN_DATA);
      // データパーンに一致しないか
      if (!matchResult) {
        // 一致しない場合

        writeContents += line_wk;

        continue;
      }

      writeContents += "\n";
      writeContents += matchResult[1] + ", " + matchResult[2] + ", " + matchResult[3];

      continue;
    }
    
    const matchResult = line_wk.match(PATTERN_DATA);
    if (!matchResult) {
      continue;
    }

    writeContents += "\n";
    writeContents += matchResult[1] + ", " + matchResult[2] + ", " + matchResult[3];
  }

  writeContents = writeContents.substring(1, writeContents.length);

  createCsvFile(folderId, fileName, writeContents);
}

/**
 * CSVファイル作成する。
 * @param {string} folderId フォルダID
 * @param {string} fileName ファイル名
 * @param {string} contents ファイルの内容
 */
function createCsvFile(folderId, fileName, contents) {  

  const contentType = 'text/csv';                     // コンテンツタイプ
  const charset = 'UTF-8';                            // 文字コード
  const folder = DriveApp.getFolderById(folderId);    // 出力するフォルダ

  const blob = Utilities.newBlob('', contentType, fileName).setDataFromString(contents, charset);
  folder.createFile(blob);
}

/**
 * メイン
 */
function main() {
  const FOLDER_ID = "<フォルダID>"

  /* ドキュメントを作成する */
  let convertedFileIds = createDocuments(FOLDER_ID);

  /* テキストを出力する */
  for (let convertedFileId of convertedFileIds) {
    let text = getText(convertedFileId);
    const fileName = DriveApp.getFileById(convertedFileId).getName() + ".csv";
    createCalendarImportFile(FOLDER_ID, fileName, text);
  }
}
