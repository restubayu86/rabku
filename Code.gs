/**
 * RAB Rumah Tangga — Google Apps Script Backend
 * Deploy sebagai Web App: Execute as "Me", Who has access: "Anyone"
 * 
 * Sheet struktur:
 *   Sheet "Items"   : id | nama | kategori | harga | jumlah | keperluan | urgensi | catatan | status | created_at | updated_at
 *   Sheet "Settings": key | value
 */

const SHEET_ID   = SpreadsheetApp.getActiveSpreadsheet().getId();
const ITEMS_SHEET    = 'Items';
const SETTINGS_SHEET = 'Settings';

// ─── CORS Headers ────────────────────────────────────────────────
function setCors(output) {
  return output
    .setMimeType(ContentService.MimeType.JSON)
    .addHeader('Access-Control-Allow-Origin', '*')
    .addHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
    .addHeader('Access-Control-Allow-Headers', 'Content-Type');
}

function ok(data)  { return setCors(ContentService.createTextOutput(JSON.stringify({ ok: true,  ...data }))); }
function err(msg)  { return setCors(ContentService.createTextOutput(JSON.stringify({ ok: false, error: msg }))); }

// ─── GET Handler ─────────────────────────────────────────────────
function doGet(e) {
  try {
    const action = (e.parameter && e.parameter.action) || 'getAll';
    switch (action) {
      case 'getAll':    return actionGetAll();
      case 'getItem':   return actionGetItem(e.parameter.id);
      case 'getBudget': return actionGetBudget();
      case 'getStats':  return actionGetStats();
      case 'ping':      return ok({ message: 'RAB API aktif', timestamp: new Date().toISOString() });
      default:          return err('Action tidak dikenal: ' + action);
    }
  } catch(ex) {
    return err('Server error: ' + ex.message);
  }
}

// ─── POST Handler ────────────────────────────────────────────────
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    switch (body.action) {
      case 'addItem':      return actionAddItem(body.item);
      case 'updateItem':   return actionUpdateItem(body.id, body.item);
      case 'updateStatus': return actionUpdateStatus(body.id, body.status);
      case 'deleteItem':   return actionDeleteItem(body.id);
      case 'setBudget':    return actionSetBudget(body.budget);
      case 'bulkAdd':      return actionBulkAdd(body.items);
      case 'deleteAll':    return actionDeleteAll();
      default:             return err('Action tidak dikenal: ' + body.action);
    }
  } catch(ex) {
    return err('Server error: ' + ex.message);
  }
}

// ─── ACTIONS ─────────────────────────────────────────────────────

function actionGetAll() {
  const sheet  = getOrCreateSheet(ITEMS_SHEET, ITEMS_HEADERS);
  const rows   = getRows(sheet);
  const budget = getBudgetValue();
  return ok({ data: rows, budget: budget, count: rows.length });
}

function actionGetItem(id) {
  if (!id) return err('id wajib diisi');
  const sheet = getOrCreateSheet(ITEMS_SHEET, ITEMS_HEADERS);
  const rows  = getRows(sheet);
  const item  = rows.find(r => r.id === id);
  if (!item) return err('Item tidak ditemukan');
  return ok({ item });
}

function actionGetBudget() {
  return ok({ budget: getBudgetValue() });
}

function actionGetStats() {
  const sheet = getOrCreateSheet(ITEMS_SHEET, ITEMS_HEADERS);
  const rows  = getRows(sheet);
  const budget = getBudgetValue();
  const aktif  = rows.filter(r => r.status !== 'Dibatalkan');
  const belum  = aktif.filter(r => r.status !== 'Sudah Dibeli');
  const sudah  = aktif.filter(r => r.status === 'Sudah Dibeli');
  
  const totalQ1 = belum.filter(r => getQuadrant(r.keperluan, r.urgensi) === 'Q1')
                        .reduce((s,r) => s + r.harga * r.jumlah, 0);
  const totalAll = belum.reduce((s,r) => s + r.harga * r.jumlah, 0);
  
  const dist = { Q1:0, Q2:0, Q3:0, Q4:0 };
  belum.forEach(r => { const q = getQuadrant(r.keperluan,r.urgensi); dist[q]++; });
  
  return ok({ stats: {
    totalAktif: belum.length, sudahDibeli: sudah.length,
    totalBiaya: totalAll, biayaQ1: totalQ1,
    sisaAnggaran: budget - totalQ1, distribusi: dist, budget
  }});
}

function actionAddItem(item) {
  if (!item || !item.nama || !item.harga) return err('Data item tidak lengkap');
  const sheet = getOrCreateSheet(ITEMS_SHEET, ITEMS_HEADERS);
  const now   = new Date().toISOString();
  const id    = 'gs_' + Date.now() + '_' + Math.random().toString(36).slice(2,6);
  const row   = [
    id, item.nama, item.kategori || 'Lainnya',
    Number(item.harga), Number(item.jumlah) || 1,
    Number(item.keperluan) || 3, Number(item.urgensi) || 3,
    item.catatan || '', item.status || 'Belum Dibeli', now, now
  ];
  sheet.appendRow(row);
  SpreadsheetApp.flush();
  return ok({ id, message: 'Item berhasil ditambahkan' });
}

function actionUpdateItem(id, item) {
  if (!id || !item) return err('id dan item wajib diisi');
  const sheet = getOrCreateSheet(ITEMS_SHEET, ITEMS_HEADERS);
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      const now = new Date().toISOString();
      sheet.getRange(i+1, 2).setValue(item.nama       || data[i][1]);
      sheet.getRange(i+1, 3).setValue(item.kategori   || data[i][2]);
      sheet.getRange(i+1, 4).setValue(Number(item.harga)     || data[i][3]);
      sheet.getRange(i+1, 5).setValue(Number(item.jumlah)    || data[i][4]);
      sheet.getRange(i+1, 6).setValue(Number(item.keperluan) || data[i][5]);
      sheet.getRange(i+1, 7).setValue(Number(item.urgensi)   || data[i][6]);
      sheet.getRange(i+1, 8).setValue(item.catatan !== undefined ? item.catatan : data[i][7]);
      sheet.getRange(i+1, 9).setValue(item.status    || data[i][8]);
      sheet.getRange(i+1, 11).setValue(now);
      SpreadsheetApp.flush();
      return ok({ message: 'Item diperbarui' });
    }
  }
  return err('Item tidak ditemukan');
}

function actionUpdateStatus(id, status) {
  if (!id || !status) return err('id dan status wajib diisi');
  const sheet = getOrCreateSheet(ITEMS_SHEET, ITEMS_HEADERS);
  const data  = sheet.getDataRange().getValues();
  const validStatus = ['Belum Dibeli', 'Sudah Dibeli', 'Dibatalkan'];
  if (!validStatus.includes(status)) return err('Status tidak valid');
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.getRange(i+1, 9).setValue(status);
      sheet.getRange(i+1, 11).setValue(new Date().toISOString());
      SpreadsheetApp.flush();
      return ok({ message: 'Status diperbarui ke: ' + status });
    }
  }
  return err('Item tidak ditemukan');
}

function actionDeleteItem(id) {
  if (!id) return err('id wajib diisi');
  const sheet = getOrCreateSheet(ITEMS_SHEET, ITEMS_HEADERS);
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.deleteRow(i + 1);
      SpreadsheetApp.flush();
      return ok({ message: 'Item dihapus' });
    }
  }
  return err('Item tidak ditemukan');
}

function actionSetBudget(budget) {
  if (!budget || isNaN(budget) || budget < 0) return err('Nilai anggaran tidak valid');
  const sheet = getOrCreateSheet(SETTINGS_SHEET, SETTINGS_HEADERS);
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === 'budget') {
      sheet.getRange(i+1, 2).setValue(Number(budget));
      SpreadsheetApp.flush();
      return ok({ budget: Number(budget), message: 'Anggaran diperbarui' });
    }
  }
  sheet.appendRow(['budget', Number(budget)]);
  SpreadsheetApp.flush();
  return ok({ budget: Number(budget), message: 'Anggaran disimpan' });
}

function actionBulkAdd(items) {
  if (!items || !Array.isArray(items)) return err('items harus berupa array');
  const sheet = getOrCreateSheet(ITEMS_SHEET, ITEMS_HEADERS);
  const ids = [];
  items.forEach(item => {
    const now = new Date().toISOString();
    const id  = 'gs_' + Date.now() + '_' + Math.random().toString(36).slice(2,6);
    sheet.appendRow([
      id, item.nama, item.kategori || 'Lainnya',
      Number(item.harga), Number(item.jumlah) || 1,
      Number(item.keperluan) || 3, Number(item.urgensi) || 3,
      item.catatan || '', item.status || 'Belum Dibeli', now, now
    ]);
    ids.push(id);
    Utilities.sleep(50);
  });
  SpreadsheetApp.flush();
  return ok({ ids, count: ids.length, message: `${ids.length} item ditambahkan` });
}

function actionDeleteAll() {
  const sheet = getOrCreateSheet(ITEMS_SHEET, ITEMS_HEADERS);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
  SpreadsheetApp.flush();
  return ok({ message: 'Semua item dihapus' });
}

// ─── UTILITIES ───────────────────────────────────────────────────

const ITEMS_HEADERS    = ['id','nama','kategori','harga','jumlah','keperluan','urgensi','catatan','status','created_at','updated_at'];
const SETTINGS_HEADERS = ['key','value'];

function getOrCreateSheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#6C63FF').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getRows(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    obj.harga     = Number(obj.harga)     || 0;
    obj.jumlah    = Number(obj.jumlah)    || 1;
    obj.keperluan = Number(obj.keperluan) || 3;
    obj.urgensi   = Number(obj.urgensi)   || 3;
    return obj;
  });
}

function getBudgetValue() {
  const sheet = getOrCreateSheet(SETTINGS_SHEET, SETTINGS_HEADERS);
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === 'budget') return Number(data[i][1]) || 5000000;
  }
  return 5000000;
}

function getQuadrant(k, u) {
  const kn = Number(k), un = Number(u);
  if (kn >= 3 && un >= 3) return 'Q1';
  if (kn >= 3 && un <  3) return 'Q2';
  if (kn <  3 && un >= 3) return 'Q3';
  return 'Q4';
}

// ─── MENU (untuk testing dari Spreadsheet) ───────────────────────
function onOpen() {
  SpreadsheetApp.getUi().createMenu('🏠 RAB App')
    .addItem('Setup Sheets', 'setupSheets')
    .addItem('Test API', 'testApi')
    .addItem('Lihat Stats', 'showStats')
    .addToUi();
}

function setupSheets() {
  getOrCreateSheet(ITEMS_SHEET, ITEMS_HEADERS);
  getOrCreateSheet(SETTINGS_SHEET, SETTINGS_HEADERS);
  SpreadsheetApp.getUi().alert('✅ Sheet berhasil dibuat/diperiksa!');
}

function testApi() {
  const res = actionGetAll();
  SpreadsheetApp.getUi().alert('API Test:\n' + res.getContent().slice(0, 300));
}

function showStats() {
  const res   = JSON.parse(actionGetStats().getContent());
  const stats = res.stats;
  SpreadsheetApp.getUi().alert(
    `📊 Statistik RAB\n\n` +
    `Total Item Aktif: ${stats.totalAktif}\n` +
    `Sudah Dibeli: ${stats.sudahDibeli}\n` +
    `Total Biaya: Rp ${stats.totalBiaya.toLocaleString('id-ID')}\n` +
    `Biaya Q1 (Mendesak): Rp ${stats.biayaQ1.toLocaleString('id-ID')}\n` +
    `Sisa Anggaran: Rp ${stats.sisaAnggaran.toLocaleString('id-ID')}\n\n` +
    `Distribusi:\n Q1: ${stats.distribusi.Q1} | Q2: ${stats.distribusi.Q2} | Q3: ${stats.distribusi.Q3} | Q4: ${stats.distribusi.Q4}`
  );
}
