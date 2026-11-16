/**
 * 讀報九宮格｜作答收件程式（Google Apps Script）
 *
 * 使用方式：
 * 1. 建立一份新的 Google 試算表 → 選單「擴充功能」→「Apps Script」
 * 2. 刪掉原本的程式碼，把這整份貼上 → 按「儲存」
 * 3. 右上角「部署」→「新增部署作業」→ 齒輪選「網頁應用程式」
 *    - 執行身分：我
 *    - 誰可以存取：任何人
 * 4. 按「部署」→ 授權 → 複製「網頁應用程式網址」（.../exec）
 *
 * 試算表會自動產生兩個工作表：
 * - 「總表」：一位學生一列，九格內容並排，方便批閱
 * - 「作答紀錄」：每次儲存都記一筆（含修改歷程）
 * 手寫圖片會存到你雲端硬碟的「讀報九宮格_手寫圖片」資料夾。
 */

const FOLDER_NAME = '讀報九宮格_手寫圖片';
const TASK_TITLES = ['標題製造機', '關鍵字 Tag', '知音相惜', '影像支援', '聲歷其境',
                     '記者 OS', '靈魂拷問', '畫重點', '生活連線'];

function doGet() {
  return ContentService.createTextOutput('讀報九宮格收件中 ✓');
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const d = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const cls = String(d.cls || '').trim();
    const seat = String(d.seat || '').trim();
    const name = String(d.name || '').trim();
    const taskId = Number(d.taskId) || 0;

    let content = String(d.content || '');
    let imageUrl = '';
    if (d.image) {
      imageUrl = saveImage_(d, cls, seat, name);
      content = imageUrl;
    }

    // 作答紀錄
    const log = getSheet_(ss, '作答紀錄', ['收到時間', '班級', '座號', '姓名', '格子', '任務', '內容']);
    log.appendRow([new Date(), cls, "'" + seat, name, taskId || '—', d.title || '', content]);

    // 總表
    const head = ['班級', '座號', '姓名', '完成格數', '已繳交 Padlet', '最後更新']
      .concat(TASK_TITLES.map((t, i) => (i + 1) + ' ' + t));
    const sum = getSheet_(ss, '總表', head);
    const row = findOrCreateRow_(sum, cls, seat, name);

    if (taskId >= 1 && taskId <= 9) {
      const cell = sum.getRange(row, 6 + taskId);
      if (imageUrl) cell.setFormula('=HYPERLINK("' + imageUrl + '","看手寫圖")');
      else cell.setValue(content);
      cell.setWrap(true);
    }
    if (taskId === 0) sum.getRange(row, 5).setValue('✓ ' + Utilities.formatDate(new Date(), 'Asia/Taipei', 'MM/dd HH:mm'));

    const vals = sum.getRange(row, 7, 1, 9).getValues()[0];
    sum.getRange(row, 4).setValue(vals.filter(v => v !== '' && v !== null).length);
    sum.getRange(row, 3).setValue(name);
    sum.getRange(row, 6).setValue(new Date());

    return ContentService.createTextOutput('ok');
  } catch (err) {
    return ContentService.createTextOutput('error: ' + err);
  } finally {
    lock.releaseLock();
  }
}

function getSheet_(ss, name, header) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(header);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, header.length).setFontWeight('bold').setBackground('#f8f4ea');
    if (name === '總表') {
      sh.setColumnWidths(7, 9, 180);
      sh.getRange('B:B').setNumberFormat('@');
    }
  }
  return sh;
}

function findOrCreateRow_(sh, cls, seat, name) {
  const last = sh.getLastRow();
  if (last >= 2) {
    const keys = sh.getRange(2, 1, last - 1, 2).getDisplayValues();
    for (let i = 0; i < keys.length; i++) {
      if (String(keys[i][0]) === cls && String(keys[i][1]) === seat) return i + 2;
    }
  }
  sh.appendRow([cls, "'" + seat, name]);
  return sh.getLastRow();
}

function saveImage_(d, cls, seat, name) {
  const it = DriveApp.getFoldersByName(FOLDER_NAME);
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
  const ext = d.mime === 'image/png' ? 'png' : 'jpg';
  const stamp = Utilities.formatDate(new Date(), 'Asia/Taipei', 'MMdd-HHmmss');
  const fname = cls + '-' + seat + '_' + name + '_' + d.title + '_' + stamp + '.' + ext;
  const blob = Utilities.newBlob(Utilities.base64Decode(d.image), d.mime || 'image/jpeg', fname);
  return folder.createFile(blob).getUrl();
}
