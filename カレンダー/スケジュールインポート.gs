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
 * メイン
 */
function main() {
  const FOLDER_ID = "<フォルダID>"

  /* ドキュメントを作成する */
  let convertedFileIds = createDocuments(FOLDER_ID);

  /* テキストを出力する */
  for (let convertedFileId of convertedFileIds) {
    let text = getText(convertedFileId);
    Logger.log(text);
  }
}
