// ============================================================
// 試作課 日報アプリ（分入力・出退勤連携版）
// ============================================================

const SHEETS = {
  CRAFTSMEN:  '職人',
  SCHEDULES:  'スケジュール',
  STAGES:     'ステージ',
  LOGS:       '日報ログ'
};

const MINUTES_PER_NINKU = 480;   // 1人工 = 8時間 = 480分
const RATE_PER_MINUTE   = 42;    // 試作分給 42円/分

// ----------------------------------------------------------
// Web アプリエントリーポイント
// ----------------------------------------------------------
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('試作課 日報')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ----------------------------------------------------------
// 初回セットアップ
// ----------------------------------------------------------
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 職人
  let s = getOrCreate(ss, SHEETS.CRAFTSMEN);
  if (s.getLastRow() === 0) {
    s.appendRow(['ID', '名前', '備考']);
    header(s, 3);
  }

  // スケジュール（製品名）
  s = getOrCreate(ss, SHEETS.SCHEDULES);
  if (s.getLastRow() === 0) {
    s.appendRow(['ID', '製品名', 'ブランド', '企画名称', 'ステータス', '備考']);
    header(s, 6);
    s.setColumnWidth(2, 220);
    s.setColumnWidth(3, 150);
    s.setColumnWidth(4, 180);
  }

  // ステージマスター
  s = getOrCreate(ss, SHEETS.STAGES);
  if (s.getLastRow() === 0) {
    s.appendRow(['ID', 'ステージ名', '順番']);
    header(s, 3);
    const defaults = [
      ['ST1', 'モック',       1],
      ['ST2', '1st',          2],
      ['ST3', '2nd',          3],
      ['ST4', '最終（展示会）', 4],
      ['ST5', '色増し前修正', 5],
      ['ST6', '試験体',       6],
      ['ST7', 'ショー用',     7],
    ];
    defaults.forEach(r => s.appendRow(r));
  }

  // 日報ログ
  s = getOrCreate(ss, SHEETS.LOGS);
  if (s.getLastRow() === 0) {
    s.appendRow([
      'ID', '日付', '職人名',
      '出勤時刻', '退勤時刻', '休憩(分)', '実働(分)',
      '製品名', 'ステージ', '作業時間(分)', '人工数', '労務費(円)',
      'メモ', '提出日時'
    ]);
    header(s, 14);
    [2,3,8,9,13].forEach(c => s.setColumnWidth(c, 120));
    s.setColumnWidth(8, 200);
  }

  return { success: true, message: 'セットアップ完了しました。' };
}

// ----------------------------------------------------------
// 初期データ取得
// ----------------------------------------------------------
function getInitialData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return {
    craftsmen: getCraftsmen(ss),
    schedules: getSchedules(ss),
    stages:    getStages(ss)
  };
}

function getCraftsmen(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  const s = ss.getSheetByName(SHEETS.CRAFTSMEN);
  if (!s || s.getLastRow() <= 1) return [];
  return s.getRange(2, 1, s.getLastRow()-1, 3).getValues()
    .filter(r => r[1]).map(r => ({ id:r[0], name:r[1], note:r[2] }));
}

function getSchedules(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  const s = ss.getSheetByName(SHEETS.SCHEDULES);
  if (!s || s.getLastRow() <= 1) return [];
  return s.getRange(2, 1, s.getLastRow()-1, 6).getValues()
    .filter(r => r[1]).map(r => ({
      id:r[0], name:r[1], brand:r[2], plan:r[3], status:r[4]||'進行中', note:r[5]
    }));
}

function getStages(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  const s = ss.getSheetByName(SHEETS.STAGES);
  if (!s || s.getLastRow() <= 1) return [];
  return s.getRange(2, 1, s.getLastRow()-1, 3).getValues()
    .filter(r => r[1])
    .sort((a,b) => a[2]-b[2])
    .map(r => ({ id:r[0], name:r[1], order:r[2] }));
}

// ----------------------------------------------------------
// 日報提出
// ----------------------------------------------------------
function submitReport(craftsmanName, dateStr, clockIn, clockOut, breakMin, rows, memo) {
  if (!craftsmanName) return { success:false, message:'職人名を選択してください。' };
  if (!dateStr)       return { success:false, message:'日付を入力してください。' };
  if (!rows || rows.length === 0) return { success:false, message:'作業行を1件以上入力してください。' };

  const actualMin  = calcActualMinutes(clockIn, clockOut, Number(breakMin)||0);
  const totalInput = rows.reduce((s,r) => s + (Number(r.minutes)||0), 0);

  const warning = (actualMin > 0 && totalInput !== actualMin)
    ? `※ 入力合計 ${totalInput}分 ／ 実働 ${actualMin}分（差: ${totalInput - actualMin}分）`
    : null;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.LOGS);
  const submittedAt = fmt(new Date(), 'yyyy/MM/dd HH:mm:ss');

  rows.forEach(row => {
    const min    = Number(row.minutes) || 0;
    const ninku  = parseFloat((min / MINUTES_PER_NINKU).toFixed(4));
    const cost   = Math.round(min * RATE_PER_MINUTE);
    sheet.appendRow([
      Utilities.getUuid(),
      dateStr, craftsmanName,
      clockIn||'', clockOut||'', Number(breakMin)||0, actualMin,
      row.productName||'', row.stageName||'',
      min, ninku, cost,
      row.memo || memo || '',
      submittedAt
    ]);
  });

  return {
    success: true,
    warning: warning,
    message: `提出しました（${rows.length}件）。${warning||''}`
  };
}

// ----------------------------------------------------------
// 実働分計算
// ----------------------------------------------------------
function calcActualMinutes(clockIn, clockOut, breakMin) {
  if (!clockIn || !clockOut) return 0;
  try {
    const [ih, im] = clockIn.split(':').map(Number);
    const [oh, om] = clockOut.split(':').map(Number);
    const total = (oh * 60 + om) - (ih * 60 + im);
    return Math.max(0, total - (breakMin||0));
  } catch(e) { return 0; }
}

// ----------------------------------------------------------
// ログ取得
// ----------------------------------------------------------
function getLogs(fCraftsman, fProduct, fStage, fFrom, fTo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s  = ss.getSheetByName(SHEETS.LOGS);
  if (!s || s.getLastRow() <= 1) return [];

  let rows = s.getRange(2, 1, s.getLastRow()-1, 14).getValues()
    .filter(r => r[0])
    .map(r => ({
      id:r[0], date:String(r[1]), craftsmanName:r[2],
      clockIn:r[3], clockOut:r[4], breakMin:r[5], actualMin:r[6],
      productName:r[7], stageName:r[8],
      minutes:r[9], ninkuCount:r[10], cost:r[11],
      memo:r[12], submittedAt:String(r[13])
    }));

  if (fCraftsman) rows = rows.filter(r => r.craftsmanName === fCraftsman);
  if (fProduct)   rows = rows.filter(r => r.productName   === fProduct);
  if (fStage)     rows = rows.filter(r => r.stageName     === fStage);
  if (fFrom)      rows = rows.filter(r => r.date >= fFrom);
  if (fTo)        rows = rows.filter(r => r.date <= fTo);

  return rows.reverse().slice(0, 300);
}

// ----------------------------------------------------------
// 集計（製品 × ステージ クロス集計）
// ----------------------------------------------------------
function getSummary(fCraftsman, fProduct, fStage, fFrom, fTo) {
  const logs   = getLogs(fCraftsman, fProduct, fStage, fFrom, fTo);
  const stages = getStages().map(s => s.name);

  // 製品ごとにステージ別に集計
  const map = {};
  logs.forEach(log => {
    const p = log.productName || '（製品名なし）';
    if (!map[p]) {
      map[p] = { productName: p, stages:{}, totalMin:0, totalCost:0 };
      stages.forEach(st => map[p].stages[st] = 0);
    }
    const m = Number(log.minutes)||0;
    if (map[p].stages[log.stageName] !== undefined) {
      map[p].stages[log.stageName] += m;
    } else {
      map[p].stages[log.stageName] = m;
    }
    map[p].totalMin  += m;
    map[p].totalCost += Number(log.cost)||0;
  });

  return {
    stages: stages,
    rows: Object.values(map)
      .sort((a,b) => b.totalMin - a.totalMin)
      .map(r => ({
        productName: r.productName,
        stages: r.stages,
        totalMin:  r.totalMin,
        totalNinku: parseFloat((r.totalMin / MINUTES_PER_NINKU).toFixed(2)),
        totalCost:  r.totalCost
      }))
  };
}

// ----------------------------------------------------------
// マスター管理
// ----------------------------------------------------------
function addCraftsman(name, note) {
  if (!name) return { success:false, message:'名前を入力してください。' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s  = ss.getSheetByName(SHEETS.CRAFTSMEN);
  s.appendRow(['C'+String(s.getLastRow()).padStart(3,'0'), name, note||'']);
  return { success:true, message:`${name} を追加しました。` };
}
function deleteCraftsman(name) { return delRow(SHEETS.CRAFTSMEN, 2, name); }

function addSchedule(name, brand, plan, status) {
  if (!name) return { success:false, message:'製品名を入力してください。' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s  = ss.getSheetByName(SHEETS.SCHEDULES);
  s.appendRow(['S'+String(s.getLastRow()).padStart(3,'0'), name, brand||'', plan||'', status||'進行中', '']);
  return { success:true, message:`「${name}」を追加しました。` };
}
function deleteSchedule(name) { return delRow(SHEETS.SCHEDULES, 2, name); }

function addStage(name) {
  if (!name) return { success:false, message:'ステージ名を入力してください。' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s  = ss.getSheetByName(SHEETS.STAGES);
  const order = s.getLastRow();
  s.appendRow(['ST'+order, name, order]);
  return { success:true, message:`ステージ「${name}」を追加しました。` };
}
function deleteStage(name) { return delRow(SHEETS.STAGES, 2, name); }

// ----------------------------------------------------------
// ユーティリティ
// ----------------------------------------------------------
function getOrCreate(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}
function header(s, n) {
  const r = s.getRange(1,1,1,n);
  r.setFontWeight('bold').setBackground('#1a73e8').setFontColor('#fff');
}
function fmt(d, pattern) {
  return Utilities.formatDate(d, 'Asia/Tokyo', pattern);
}
function delRow(sheetName, col, val) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s  = ss.getSheetByName(sheetName);
  if (!s || s.getLastRow() <= 1) return { success:false, message:'データがありません。' };
  const data = s.getRange(2, col, s.getLastRow()-1, 1).getValues();
  const idx  = data.findIndex(r => r[0] === val);
  if (idx === -1) return { success:false, message:`「${val}」が見つかりません。` };
  s.deleteRow(idx + 2);
  return { success:true, message:`「${val}」を削除しました。` };
}
