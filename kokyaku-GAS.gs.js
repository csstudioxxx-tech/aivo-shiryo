/***********************************************************************
 * Ai-voyage 顧客マップ（kokyaku.html）バックエンド  —  汎用KVストア
 * concurrent-safe: 1レコード=1行 / LockService / updatedAtガード
 *
 * ◆ セットアップ（5分）
 *  1) スプレッドシートを新規作成（名前は任意。例「Ai-voyage 顧客DB」）。
 *  2) そのスプレッドシートの URL から ID をコピー
 *     （https://docs.google.com/spreadsheets/d/【ここがID】/edit）
 *  3) 拡張機能 → Apps Script を開き、このコードを全文貼り付け。
 *  4) 上のメニュー「プロジェクトの設定（歯車）」→「スクリプト プロパティ」
 *     →  プロパティ名: SSID   値: 手順2のID  を追加して保存。
 *  5) 「デプロイ」→「新しいデプロイ」→種類「ウェブアプリ」
 *       - 実行するユーザー: 自分
 *       - アクセスできるユーザー: 全員
 *     →デプロイ →表示される「ウェブアプリ URL（/exec で終わる）」をコピー。
 *  6) kokyaku.html の先頭付近 CFG.GAS_URL にそのURLを貼り付け、再アップロード。
 *  7) 全員がページをスーパーリロード（Ctrl/Cmd+Shift+R）。
 *
 * ◆ コードを更新したとき
 *  「デプロイ」→「デプロイを管理」→鉛筆→バージョン「新バージョン」→デプロイ。
 *  （URLは変わりません）
 ***********************************************************************/

var SHEET_NAME = 'kv';

function _sheet_() {
  var ssid = PropertiesService.getScriptProperties().getProperty('SSID');
  if (!ssid) throw new Error('Script Property "SSID" が未設定です。');
  var ss = SpreadsheetApp.openById(ssid);
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, 3).setValues([['key', 'value', 'updatedAt']]);
  }
  return sh;
}

function _json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** 全KVを返す: {ok:true, data:{key:value(JSON文字列)}} */
function doGet(e) {
  try {
    var sh = _sheet_();
    var last = sh.getLastRow();
    var data = {};
    if (last >= 2) {
      var rows = sh.getRange(2, 1, last - 1, 2).getValues();
      for (var i = 0; i < rows.length; i++) {
        var k = rows[i][0];
        if (k) data[k] = rows[i][1];
      }
    }
    return _json_({ ok: true, data: data });
  } catch (err) {
    return _json_({ ok: false, error: String(err) });
  }
}

/** 1レコード upsert: body {action:"set", key, value, updatedAt}
 *  updatedAt ガード: 既存より古い/欠損の書き込みは無視（=上書き事故防止） */
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (err) {
    return _json_({ ok: false, error: 'busy' });
  }
  try {
    var body = {};
    if (e && e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    }
    if (body.action !== 'set' || !body.key) {
      return _json_({ ok: false, error: 'bad request' });
    }
    var sh = _sheet_();
    var last = sh.getLastRow();
    var incomingUpdated = Number(body.updatedAt || 0);

    var foundRow = -1, existingUpdated = -1;
    if (last >= 2) {
      var keys = sh.getRange(2, 1, last - 1, 1).getValues();
      var ups = sh.getRange(2, 3, last - 1, 1).getValues();
      for (var i = 0; i < keys.length; i++) {
        if (keys[i][0] === body.key) {
          foundRow = i + 2;
          existingUpdated = Number(ups[i][0] || 0);
          break;
        }
      }
    }

    // 既存が同じか新しければ無視（古い書き込みで他者の更新を消さない）
    if (foundRow > 0 && existingUpdated > incomingUpdated) {
      return _json_({ ok: true, skipped: true, reason: 'stale' });
    }

    if (foundRow > 0) {
      sh.getRange(foundRow, 2, 1, 2).setValues([[String(body.value || ''), incomingUpdated]]);
    } else {
      sh.appendRow([body.key, String(body.value || ''), incomingUpdated]);
    }
    return _json_({ ok: true });
  } catch (err) {
    return _json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
