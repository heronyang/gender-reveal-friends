/**
 * 多多性別投票（朋友版）— Google Apps Script 後端
 *
 * 這是「獨立」的 Apps Script 專案（從 script.google.com 新增，不是從試算表的「擴充功能」開）。
 * 它和家人版共用同一個試算表，但只讀寫「朋友投票」分頁，不會碰到家人的「投票」分頁。
 *
 * 試算表 ID 不寫在程式碼裡（repo 是公開的），請放在：
 *   專案設定（齒輪）→ 指令碼屬性 → 新增 SPREADSHEET_ID = 試算表網址中 /d/ 和 /edit 之間那一串
 *
 * 部署：部署 → 新增部署作業 → 網頁應用程式；執行身分：我；誰可以存取：所有人
 */

const SHEET_NAME = '朋友投票';
const HEADERS = ['時間', '姓名', '猜測', '理由'];
const GUESS_LABEL = { boy: '男生', girl: '女生' };

function getSheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('尚未設定 SPREADSHEET_ID（專案設定 → 指令碼屬性）');
  const ss = SpreadsheetApp.openById(id);
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
  }
  return sh;
}

// 避免試算表公式注入（例如名字以 = 開頭）
function safe_(s) {
  s = String(s == null ? '' : s).trim();
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function readVotes_() {
  const values = getSheet_().getDataRange().getValues();
  const votes = [];
  for (let i = 1; i < values.length; i++) {
    const [time, name, guess, reason] = values[i];
    if (!name) continue;
    votes.push({
      time: time instanceof Date ? time.toISOString() : String(time),
      name: String(name).replace(/^'/, ''),
      // 也接受舊版的「男寶／女寶」
      guess: /^男/.test(guess) ? 'boy' : /^女/.test(guess) ? 'girl' : String(guess),
      reason: String(reason).replace(/^'/, ''),
    });
  }
  return votes;
}

function doGet() {
  return json_({ ok: true, votes: readVotes_() });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const name = String(d.name || '').trim().slice(0, 40);
    const reason = String(d.reason || '').trim().slice(0, 500);
    const guess = GUESS_LABEL[d.guess];
    if (!name) return json_({ ok: false, error: '請填寫姓名' });
    if (!guess) return json_({ ok: false, error: '請選擇男生或女生' });
    if (!reason) return json_({ ok: false, error: '請寫下你的理由' });

    const sh = getSheet_();
    const names = sh.getDataRange().getValues().map(r => String(r[1]).replace(/^'/, '').trim());
    const row = [new Date(), safe_(name), guess, safe_(reason)];
    const idx = names.indexOf(name, 1);
    if (idx > 0) {
      // 同名者重新投票：覆蓋原本的票，一人只算一票
      sh.getRange(idx + 1, 1, 1, row.length).setValues([row]);
      return json_({ ok: true, updated: true });
    }
    sh.appendRow(row);
    return json_({ ok: true, updated: false });
  } catch (err) {
    return json_({ ok: false, error: '伺服器錯誤：' + err });
  } finally {
    lock.releaseLock();
  }
}
