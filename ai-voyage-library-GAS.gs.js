/************************************************************************
 * Ai-voyage 説明資料ライブラリ｜共有メモ・タグ用 GAS ウェブアプリ
 *   ★ 同時編集に強い版（6名が同時に編集してもデータを消さない）
 *
 * 仕組み：
 *  - データは「1メモ＝1行」「1タグ＝1行」のレコード単位（key→value）。
 *  - 保存は該当レコード1件だけを upsert（他人のレコードは触らない）。
 *  - LockService で直列化。シートは絶対に clear（全消し）しない。
 *  - 各レコードの updatedAt(ms) を比較し、古い保存が新しいレコードを上書きしない。
 *
 * 【セットアップ手順】
 * 1) https://script.google.com →「新しいプロジェクト」
 * 2) 既定のコードを全部消して、このファイルの中身を丸ごと貼り付け → 保存
 * 3) 「デプロイ」→「新しいデプロイ」→ 種類：ウェブアプリ
 *      実行するユーザー：自分／アクセスできるユーザー：全員 → デプロイ →（承認）→ /exec URL をコピー
 * 4) index.html の  GAS_URL:""  にURLを貼る
 *
 * 【コードを直したとき（重要）】
 *   「デプロイを管理」→ 既存デプロイの鉛筆 → バージョン「新しいバージョン」→ デプロイ（URLは変わりません）。
 ************************************************************************/

var SHEET_NAME = 'kv';

function getSheet_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SSID');
  var ss = null;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { id = null; } }
  if (!ss) {
    ss = SpreadsheetApp.create('Ai-voyage 資料ライブラリ DB');
    props.setProperty('SSID', ss.getId());
  }
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, 2).setValues([['key', 'value']]);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function upd_(jsonStr) {
  try { var o = JSON.parse(jsonStr); return Number(o && o.upd) || 0; } catch (e) { return 0; }
}

// 読み込み：全レコード → { ok:true, data:{ key:value(JSON文字列), ... } }
function doGet(e) {
  try {
    var sh = getSheet_();
    var vals = sh.getDataRange().getValues();
    var data = {};
    for (var i = 1; i < vals.length; i++) {
      if (vals[i][0]) data[vals[i][0]] = vals[i][1];
    }
    return json_({ ok: true, data: data });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

// 書き込み：{ action:"set", key, value } を1レコードだけ upsert（updatedAt ガード付き）
// ※ 全消し・全置換は行わない。
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var req = JSON.parse(e.postData.contents);
    if (req.action === 'set' && req.key != null) {
      var sh = getSheet_();
      var vals = sh.getDataRange().getValues();
      var row = -1;
      for (var i = 1; i < vals.length; i++) {
        if (vals[i][0] === req.key) { row = i + 1; break; }
      }
      if (row > 0) {
        var incoming = upd_(req.value);
        var existing = upd_(vals[row - 1][1]);
        // 古い保存(または updatedAt 未設定=0)が、新しいレコードを上書きしないようにする
        if (existing > 0 && incoming < existing) {
          return json_({ ok: true, skipped: true });
        }
        sh.getRange(row, 2).setValue(req.value);
      } else {
        sh.appendRow([req.key, req.value]);
      }
      return json_({ ok: true });
    }
    return json_({ ok: false, error: 'unknown action' });
  } finally {
    lock.releaseLock();
  }
}
