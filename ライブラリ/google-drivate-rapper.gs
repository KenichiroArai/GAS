/**
 * 指定されたパスに基づいてファイルまたはフォルダのIDを取得します。
 *
 * @param {string} path - ファイルまたはフォルダのパスを '/' で区切った文字列。
 * @returns {string|undefined} - 指定されたパスに対応するファイルまたはフォルダのID。見つからない場合は undefined。
 */
function getItemIdFromPath(path) {

    // 結果を初期化
    let result = undefined;

    // ルートフォルダから開始
    let targetFolder = DriveApp.getRootFolder();

    // パスを個々のフォルダ名とファイル名に分割
    let pathItems = path.split('/');

    // パス内の各アイテムを順に処理
    for (let pathItem of pathItems) {

        // パスアイテムのUnicode正規化を行い、文字コードを統一
        pathItem = pathItem.normalize('NFC');

        // 現在のパスアイテムのIDを格納する変数
        let targetId = undefined;

        // 現在のフォルダ内のサブフォルダを確認
        let folders = DriveApp.getFolderById(targetFolder.getId()).getFolders();
        while (folders.hasNext()) {
            const folder = folders.next();
            if (folder.getName() !== pathItem) {
                // フォルダ名が一致しない場合は次のフォルダに進む
                continue;
            }
            targetId = folder.getId();
            break;
        }

        // サブフォルダに一致するものが見つからないか
        if (targetId == undefined) {
            // 見つからない場合、ファイルを検索

            let files = DriveApp.getFolderById(targetFolder.getId()).getFiles();
            while (files.hasNext()) {
                let file = files.next();
                if (file.getName() !== pathItem) {
                    // ファイル名が一致しない場合は次のファイルに進む
                    continue;
                }

                // ショートカットの場合は「file#getTargetId()」で取得できる
                targetId = file.getTargetId();
                if (targetId == null) {
                    // 通常のファイルの場合は、「file#getId()」で取得できる
                    targetId = file.getId();
                }
                break;
            }
        }

        // IDが取得できなかったか
        if (targetId === undefined) {
            // 出来なかった場合
            // 処理を終了し、undefinedを返す（パスが無効）
            result = undefined;
            return result;
        }

        // 見つかったフォルダまたはファイルを次のターゲットフォルダとして設定
        targetFolder = DriveApp.getFolderById(targetId);
        result = targetId; // 最後に見つかったターゲットIDを保存
    }

    // 最後に見つかったファイルまたはフォルダのIDを返す
    return result;
}
