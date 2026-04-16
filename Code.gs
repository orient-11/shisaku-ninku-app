// ============================================================
// 試作課 日報型 人工管理アプリ - Google Apps Script
// ============================================================

const SHEETS = {
  CRAFTSMEN:  '職人',
  SCHEDULES:  'スケジュール',
  PROCESSES:  '工程',
  DAILY_LOGS: '日報ログ'
};

const WORK_HOURS_PER_DAY = 8;  // 1人工 = 8時間

// ----------------------------------------------------------
// Web アプリエントリーポイント
// ----------------------------------------------------------
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('試作課 日報')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ----------------------------------------------------------
// 初回セットアップ
// ----------------------------------------------------------
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 職人シート
  let sheet = getOrCreateSheet(ss, SHEETS.CRAFTSMEN);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['ID', '名前', '備考']);
    styleHeader(sheet, 3);
    sheet.setColumnWidth(1, 60);
    sheet.setColumnWidth(2, 150);
    sheet.setColumnWidth(3, 200);
  }

  // スケジュールシート
  sheet = getOrCreateSheet(ss, SHEETS.SCHEDULES);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['ID', '案件名', '開始日', '終了日', 'ステータス', '備考']);
    styleHeader(sheet, 6);
    sheet.setColumnWidth(2, 220);
  }

  // 工程マスターシート
  sheet = getOrCreateSheet(ss, SHEETS.PROCESSES);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['ID', '工程名', '備考']);
    styleHeader(sheet, 3);
    // デフォルト工程を挿入
    const defaults = [
      ['P001', 'パターン作成', ''],
      ['P002', '仮縫い', ''],
      ['P003', '本縫い', ''],
      ['P004', '裁断', ''],
      ['P005', '仕上げ', ''],
      ['P006', '検品・確認', ''],
      ['P007', '打合せ・修正', ''],
      ['P008', 'その他', '']
    ];
    defaults.forEach(row => sheet.appendRow(row));
    sheet.setColumnWidth(2, 160);
  }

  // 日報ログシート
  sheet = getOrCreateSheet(ss, SHEETS.DAILY_LOGS);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'ID', '日付', '職人名', '案件名', '工程',
      '割合(%)', '時間数(h)', '人工数', 'メモ', '提出日時'
    ]);
    styleHeader(sheet, 10);
    sheet.setColumnWidth(2, 100);
    sheet.setColumnWidth(3, 120);
    sheet.setColumnWidth(4, 200);
    sheet.setColumnWidth(5, 120);
    sheet.setColumnWidth(6, 80);
    sheet.setColumnWidth(7, 80);
    sheet.setColumnWidth(8, 80);
    sheet.setColumnWidth(9, 200);
    sheet.setColumnWidth(10, 150);
  }

  return { success: true, message: 'セットアップが完了しました。' };
}

// ----------------------------------------------------------
// 初期データ取得
// ----------------------------------------------------------
function getInitialData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return {
    craftsmen:  getCraftsmen(ss),
    schedules:  getSchedules(ss),
    processes:  getProcesses(ss)
  };
}

function getCraftsmen(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.CRAFTSMEN);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 3)
    .getValues().filter(r => r[1])
    .map(r => ({ id: r[0], name: r[1], note: r[2] }));
}

function getSchedules(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.SCHEDULES);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 6)
    .getValues().filter(r => r[1])
    .map(r => ({
      id: r[0], name: r[1],
      startDate: r[2] ? Utilities.formatDate(new Date(r[2]), 'Asia/Tokyo', 'yyyy/MM/dd') : '',
      endDate:   r[3] ? Utilities.formatDate(new Date(r[3]), 'Asia/Tokyo', 'yyyy/MM/dd') : '',
      status:    r[4] || '進行中',
      note:      r[5]
    }));
}

function getProcesses(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.PROCESSES);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 3)
    .getValues().filter(r => r[1])
    .map(r => ({ id: r[0], name: r[1], note: r[2] }));
}

// ----------------------------------------------------------
// 日報提出
// ----------------------------------------------------------
function submitDailyReport(craftsmanName, dateStr, rows, dailyMemo) {
  if (!craftsmanName) return { success: false, message: '職人名を選択してください。' };
  if (!dateStr)       return { success: false, message: '日付を入力してください。' };
  if (!rows || rows.length === 0) return { success: false, message: '作業行を1件以上追加してください。' };

  // 割合の合計チェック（警告のみ、保存はする）
  const totalPct = rows.reduce((s, r) => s + (Number(r.pct) || 0), 0);
  const warning = (totalPct !== 100)
    ? `※ 割合の合計が ${totalPct}% です（100%推奨）。`
    : null;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.DAILY_LOGS);
  const submittedAt = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');

  rows.forEach(row => {
    const pct     = Number(row.pct) || 0;
    const hours   = parseFloat((pct / 100 * WORK_HOURS_PER_DAY).toFixed(2));
    const ninkuCount = parseFloat((hours / WORK_HOURS_PER_DAY).toFixed(3));
    const logId   = Utilities.getUuid();

    sheet.appendRow([
      logId,
      dateStr,
      craftsmanName,
      row.scheduleName || '（案件なし）',
      row.processName  || '',
      pct,
      hours,
      ninkuCount,
      row.memo || dailyMemo || '',
      submittedAt
    ]);
  });

  return {
    success: true,
    warning: warning,
    message: `${craftsmanName} の日報を提出しました（${rows.length}件）。${warning || ''}`
  };
}

// ----------------------------------------------------------
// 日報ログ取得
// ----------------------------------------------------------
function getLogs(filterCraftsman, filterSchedule, filterProcess, filterDateFrom, filterDateTo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.DAILY_LOGS);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  let logs = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10)
    .getValues().filter(r => r[0])
    .map(r => ({
      id: r[0], date: r[1], craftsmanName: r[2], scheduleName: r[3],
      processName: r[4], pct: r[5], hours: r[6], ninkuCount: r[7],
      memo: r[8], submittedAt: r[9]
    }));

  if (filterCraftsman) logs = logs.filter(l => l.craftsmanName === filterCraftsman);
  if (filterSchedule)  logs = logs.filter(l => l.scheduleName  === filterSchedule);
  if (filterProcess)   logs = logs.filter(l => l.processName   === filterProcess);
  if (filterDateFrom)  logs = logs.filter(l => String(l.date)  >= filterDateFrom);
  if (filterDateTo)    logs = logs.filter(l => String(l.date)  <= filterDateTo);

  return logs.reverse().slice(0, 300);
}

// ----------------------------------------------------------
// 集計取得
// ----------------------------------------------------------
function getSummary(filterCraftsman, filterSchedule, filterProcess, filterDateFrom, filterDateTo) {
  const logs = getLogs(filterCraftsman, filterSchedule, filterProcess, filterDateFrom, filterDateTo);

  // 職人 × 案件 × 工程 でグループ集計
  const map = {};
  logs.forEach(log => {
    const key = `${log.craftsmanName}||${log.scheduleName}||${log.processName}`;
    if (!map[key]) {
      map[key] = {
        craftsmanName: log.craftsmanName,
        scheduleName:  log.scheduleName,
        processName:   log.processName,
        totalHours:    0,
        totalNinku:    0,
        count:         0
      };
    }
    map[key].totalHours += Number(log.hours) || 0;
    map[key].count++;
  });

  return Object.values(map).map(s => ({
    ...s,
    totalHours:  parseFloat(s.totalHours.toFixed(2)),
    totalNinku:  parseFloat((s.totalHours / WORK_HOURS_PER_DAY).toFixed(3))
  })).sort((a, b) => b.totalHours - a.totalHours);
}

// ----------------------------------------------------------
// マスター管理
// ----------------------------------------------------------
function addCraftsman(name, note) {
  if (!name) return { success: false, message: '名前を入力してください。' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.CRAFTSMEN);
  const id = 'C' + String(sheet.getLastRow()).padStart(3, '0');
  sheet.appendRow([id, name, note || '']);
  return { success: true, message: `${name} を追加しました。` };
}

function deleteCraftsman(name) {
  return deleteRowByColumn(SHEETS.CRAFTSMEN, 2, name, '職人');
}

function addSchedule(name, startDate, endDate, status, note) {
  if (!name) return { success: false, message: '案件名を入力してください。' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.SCHEDULES);
  const id = 'S' + String(sheet.getLastRow()).padStart(3, '0');
  sheet.appendRow([id, name, startDate || '', endDate || '', status || '進行中', note || '']);
  return { success: true, message: `「${name}」を追加しました。` };
}

function deleteSchedule(name) {
  return deleteRowByColumn(SHEETS.SCHEDULES, 2, name, '案件');
}

function addProcess(name, note) {
  if (!name) return { success: false, message: '工程名を入力してください。' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.PROCESSES);
  const id = 'P' + String(sheet.getLastRow()).padStart(3, '0');
  sheet.appendRow([id, name, note || '']);
  return { success: true, message: `工程「${name}」を追加しました。` };
}

function deleteProcess(name) {
  return deleteRowByColumn(SHEETS.PROCESSES, 2, name, '工程');
}

// ----------------------------------------------------------
// ユーティリティ
// ----------------------------------------------------------
function getOrCreateSheet(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function styleHeader(sheet, colCount) {
  const range = sheet.getRange(1, 1, 1, colCount);
  range.setFontWeight('bold');
  range.setBackground('#1a73e8');
  range.setFontColor('#ffffff');
}

function deleteRowByColumn(sheetName, colIndex, value, label) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() <= 1) return { success: false, message: 'データがありません。' };
  const data = sheet.getRange(2, colIndex, sheet.getLastRow() - 1, 1).getValues();
  const idx  = data.findIndex(r => r[0] === value);
  if (idx === -1) return { success: false, message: `${label}「${value}」が見つかりません。` };
  sheet.deleteRow(idx + 2);
  return { success: true, message: `${label}「${value}」を削除しました。` };
}
