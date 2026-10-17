/**
 * 手書き一問一答アプリ → このスプレッドシートに学習記録を書き込むスクリプト
 *
 * 使い方
 *  1. 新しいGoogleスプレッドシートを作る
 *  2. メニュー「拡張機能」→「Apps Script」を開き、最初からある内容を消してこのコードを貼り付けて保存
 *  3. 右上「デプロイ」→「新しいデプロイ」→ 種類の選択（歯車）で「ウェブアプリ」
 *       次のユーザーとして実行：自分 ／ アクセスできるユーザー：全員
 *  4. 「デプロイ」→ Googleアカウントの承認 → 表示された「ウェブアプリのURL」をコピー
 *  5. アプリの ⚙設定「記録の送信先」に貼り付けて保存 →「テスト送信」
 *
 * 合言葉を使う場合は、下の TOKEN に好きな文字を入れて保存し、
 * 「デプロイを管理」→ 編集（鉛筆）→ バージョン「新バージョン」でデプロイし直してから、
 * アプリの設定にも同じ合言葉を入れてください。
 */
const TOKEN = "";  // 合言葉（空欄なら確認しない）

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (TOKEN && d.token !== TOKEN) return out_({ ok: false, error: "token" });

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const name = String(d.sheet || "記録");
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    const row = d.row || {};

    // 見出し行（足りない項目は右に追加）
    const header = sh.getLastRow() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String) : [];
    let changed = false;
    Object.keys(row).forEach(k => { if (header.indexOf(k) < 0) { header.push(k); changed = true; } });
    if (changed) {
      sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight("bold").setBackground("#efe9dc");
      sh.setFrozenRows(1);
    }
    const values = header.map(h => (h in row ? row[h] : ""));

    // 1列目（日付・週）が同じ行があれば上書き、なければ下に追加
    const last = sh.getLastRow();
    let target = 0;
    if (last >= 2) {
      const keys = sh.getRange(2, 1, last - 1, 1).getDisplayValues().map(r => r[0]);
      const i = keys.indexOf(String(d.key));
      if (i >= 0) target = i + 2;
    }
    if (!target) target = last + 1;
    sh.getRange(target, 1).setNumberFormat("@");   // 日付が自動で変換されないよう文字列で保存
    sh.getRange(target, 1, 1, values.length).setValues([values]).setVerticalAlignment("top");
    return out_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
