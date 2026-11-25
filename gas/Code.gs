/**
 * 瓶子點數助手：接收盤點紀錄，寫入 Google 試算表並把標記照片存進 Drive。
 * 第一次使用：在編輯器執行 setup()，授權後會在共用資料夾裡建立試算表與「標記照片」資料夾。
 */
const P = PropertiesService.getScriptProperties();
const TARGET_FOLDER_ID = '1kLNVYhdmkGdBHD290hOU0x0vYPDCYDSl'; // 共用的盤點資料夾
const HEADER = ['存檔時間', '箱號／品名', '盤點人', '總數', '照片張數', '明細', '照片資料夾', '照片連結'];

function setup() {
  if (!P.getProperty('SHEET_ID')) {
    const folder = DriveApp.getFolderById(TARGET_FOLDER_ID);
    const photos = folder.createFolder('標記照片');
    const ss = SpreadsheetApp.create('瓶子盤點紀錄');
    DriveApp.getFileById(ss.getId()).moveTo(folder);
    const sh = ss.getSheets()[0];
    sh.setName('盤點紀錄');
    sh.getRange(1, 1, 1, HEADER.length).setValues([HEADER])
      .setFontWeight('bold').setHorizontalAlignment('center')
      .setBackground('#215F9A').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    sh.setColumnWidths(1, HEADER.length, 140);
    sh.setColumnWidth(6, 320);
    sh.getRange('D:E').setNumberFormat('#,##0');
    P.setProperties({ FOLDER_ID: folder.getId(), PHOTO_ID: photos.getId(), SHEET_ID: ss.getId() });
  }
  // 程式本身也放進共用資料夾，方便管理
  const self = DriveApp.getFileById(ScriptApp.getScriptId());
  const target = DriveApp.getFolderById(TARGET_FOLDER_ID);
  if (!self.getParents().hasNext() || self.getParents().next().getId() !== target.getId()) self.moveTo(target);
  const url = SpreadsheetApp.openById(P.getProperty('SHEET_ID')).getUrl();
  Logger.log('試算表：' + url);
  Logger.log('資料夾：' + DriveApp.getFolderById(P.getProperty('FOLDER_ID')).getUrl());
  return url;
}

function doGet() {
  return json_({ ok: true, ready: !!P.getProperty('SHEET_ID') });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
    const d = JSON.parse(e.postData.contents);
    if (!P.getProperty('SHEET_ID')) return json_({ ok: false, error: '尚未完成 setup' });
    if (d.action === 'ping') return json_({ ok: true, ready: true });

    const box = clean_(d.box, 60), person = clean_(d.person, 30);
    if (!box || !person) return json_({ ok: false, error: '請填箱號／品名與盤點人' });
    const entries = Array.isArray(d.entries) ? d.entries.slice(0, 50) : [];
    if (!entries.length) return json_({ ok: false, error: '沒有點數記錄' });

    // 總數由伺服器重算，與畫面上確認的總數不一致就不存
    let total = 0;
    const lines = entries.map((en, i) => {
      const n = Math.max(0, Math.floor(Number(en.n) || 0));
      const m = Math.max(1, Math.floor(Number(en.mult) || 1));
      total += n * m;
      return '第' + (i + 1) + '張：' + n + ' × ' + m + ' = ' + n * m;
    });
    if (Number(d.total) !== total) return json_({ ok: false, error: '總數核對不一致（伺服器算出 ' + total + '），請重新整理後再試' });

    const now = new Date();
    const stamp = Utilities.formatDate(now, 'Asia/Taipei', 'yyyyMMdd-HHmmss');
    const photoFolder = DriveApp.getFolderById(P.getProperty('PHOTO_ID'));
    const sub = photoFolder.createFolder(stamp + '_' + box.replace(/[\\/:*?"<>|]/g, '_'));
    const links = [];
    entries.forEach((en, i) => {
      const m = /^data:image\/jpeg;base64,(.+)$/.exec(String(en.img || ''));
      if (!m) return;
      const f = sub.createFile(Utilities.newBlob(Utilities.base64Decode(m[1]), 'image/jpeg', '第' + (i + 1) + '張.jpg'));
      links.push(f.getUrl());
    });

    const sh = SpreadsheetApp.openById(P.getProperty('SHEET_ID')).getSheets()[0];
    sh.appendRow([now, box, person, total, links.length, lines.join('\n'), sub.getUrl(), links.join('\n')]);
    const row = sh.getLastRow();
    sh.getRange(row, 1).setNumberFormat('yyyy/mm/dd hh:mm');
    return json_({ ok: true, total: total, row: row, photos: links.length });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

function clean_(s, max) {
  s = String(s == null ? '' : s).replace(/[\r\n\t]+/g, ' ').trim().slice(0, max);
  return /^[=+\-@]/.test(s) ? "'" + s : s; // 防止被試算表當成公式
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
