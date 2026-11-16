/**
 * 金拾煎 線上訂單後台（Google Apps Script，綁定在 Google 試算表上）
 *
 * 功能
 *  1. doPost   ：接收網站結帳送來的訂單 → 依「菜單」工作表重新計價 → 寫入「訂單」「訂單明細」
 *  2. Email    ：新訂單通知店家；客人有填 Email 時寄出訂單確認信
 *  3. 狀態通知 ：在「訂單」把狀態改成「可取餐／已出貨／已取消」時，自動寄信通知客人
 *  4. 營業日報 ：每天晚上寄出當日訂單數、營業額、熱銷品項
 *  5. doGet    ：健康檢查，以及以「訂單編號＋電話」查詢訂單狀態
 *
 * 第一次使用：在 Apps Script 編輯器執行 setup()（只要一次），再「部署 → 網頁應用程式」。
 * 詳細步驟見 gas/README.md。
 */

/* =========================== 設定 =========================== */
const CONFIG = {
  SHOP_NAME: '金拾煎',
  SPREADSHEET_ID: '',        // 試算表 ID（網址 /d/ 與 /edit 中間那串）；從試算表「擴充功能 → Apps Script」建立的可留空
  NOTIFY_EMAIL: '',          // 收新訂單通知與日報的信箱；留空＝部署這個腳本的 Google 帳號
  TIMEZONE: 'Asia/Taipei',
  DELIVERY_FEE: 150,         // 冷凍宅配運費
  FREE_SHIP: 1500,           // 滿額免運門檻
  DINE_RATE: 0.1,            // 內用服務費比例
  DAILY_REPORT_HOUR: 21,     // 日報寄送時間（21 點左右）
  MAX_ORDERS_PER_PHONE: 5,   // 同一支電話 10 分鐘內最多下單次數（防灌單）
  STORES: { flagship: '旗艦店', xinyi: '信義店', banqiao: '板橋店' }
};

const SHEET = { ORDERS: '訂單', ITEMS: '訂單明細', MENU: '菜單', LOG: '錯誤紀錄' };
const STATUSES = ['新訂單', '製作中', '可取餐', '已出貨', '已完成', '已取消'];
const MODE_TEXT = { dine: '內用', takeout: '外帶', delivery: '冷凍宅配' };

const ORDER_HEADERS = ['訂單編號', '建立時間', '狀態', '取餐方式', '門市', '取餐時間', '姓名', '電話', 'Email',
  '地址', '備註', '品項摘要', '小計', '服務費/運費', '總額', '通知紀錄'];
const COL = ORDER_HEADERS.reduce((m, h, i) => (m[h] = i + 1, m), {});   // 欄名 → 欄號（從 1 開始）
const ITEM_HEADERS = ['訂單編號', '建立時間', '品項代碼', '品項名稱', '單價', '數量', '單位', '金額'];
const MENU_HEADERS = ['品項代碼', '品項名稱', '分類', '單價', '單位', '最小單位', '上架'];

/* 菜單初始資料（與網站一致）；之後改價、下架請直接改「菜單」工作表 */
const MENU_SEED = [
  ['p1', '招牌高麗菜鍋貼', '招牌鍋貼', 6, '顆', 10], ['p2', '韭菜豬肉鍋貼', '招牌鍋貼', 6, '顆', 10],
  ['p3', '麻辣起司鍋貼', '招牌鍋貼', 8, '顆', 10], ['p4', '蝦仁鮮肉鍋貼', '招牌鍋貼', 10, '顆', 10],
  ['d1', '鮮肉水餃', '精選水餃', 5, '顆', 10], ['d2', '玉米鮮肉水餃', '精選水餃', 6, '顆', 10],
  ['d3', '紅油抄手', '精選水餃', 75, '份', 1], ['d4', '蔬食香菇水餃', '精選水餃', 6, '顆', 10],
  ['s1', '酸辣湯', '湯品/飲料', 45, '碗', 1], ['s2', '玉米濃湯', '湯品/飲料', 40, '碗', 1],
  ['s3', '古早味紅茶', '湯品/飲料', 30, '杯', 1], ['s4', '無糖豆漿', '湯品/飲料', 35, '杯', 1],
  ['c1', '涼拌小黃瓜', '經典小菜', 40, '份', 1], ['c2', '滷味拼盤', '經典小菜', 60, '份', 1],
  ['c3', '皮蛋豆腐', '經典小菜', 45, '份', 1], ['c4', '辣味干絲', '經典小菜', 45, '份', 1],
  ['g1', '招牌招財鍋貼禮盒', '冷凍禮盒', 520, '盒', 1], ['g2', '家庭常備水餃組', '冷凍禮盒', 680, '盒', 1],
  ['g3', '節慶金饌伴手禮盒', '冷凍禮盒', 880, '盒', 1]
].map((r) => r.concat('Y'));

/* =========================== 初始化 =========================== */
/** 建立工作表、標題、格式、下拉選單，並安裝觸發器。可重複執行，不會刪資料。 */
function setup() {
  const ss = ss_();
  ss.setSpreadsheetTimeZone(CONFIG.TIMEZONE);

  const orders = ensureSheet_(ss, SHEET.ORDERS, ORDER_HEADERS);
  orders.getRange(2, COL['訂單編號'], orders.getMaxRows() - 1, 1).setNumberFormat('@');
  orders.getRange(2, COL['電話'], orders.getMaxRows() - 1, 1).setNumberFormat('@');
  orders.getRange(2, COL['建立時間'], orders.getMaxRows() - 1, 1).setNumberFormat('yyyy/mm/dd hh:mm');
  orders.getRange(2, COL['小計'], orders.getMaxRows() - 1, 3).setNumberFormat('#,##0');
  const statusRange = orders.getRange(2, COL['狀態'], orders.getMaxRows() - 1, 1);
  statusRange.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).setAllowInvalid(false).build());
  const colors = { '新訂單': '#fde2cf', '製作中': '#fff2b3', '可取餐': '#d4f0d4', '已出貨': '#d3e6fb', '已完成': '#e6e6e6', '已取消': '#f6c9c9' };
  orders.setConditionalFormatRules(Object.keys(colors).map((s) =>
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(s).setBackground(colors[s]).setRanges([statusRange]).build()));
  [130, 130, 80, 80, 80, 100, 90, 110, 180, 220, 180, 260, 70, 90, 80, 220].forEach((w, i) => orders.setColumnWidth(i + 1, w));

  const items = ensureSheet_(ss, SHEET.ITEMS, ITEM_HEADERS);
  items.getRange(2, 1, items.getMaxRows() - 1, 1).setNumberFormat('@');
  items.getRange(2, 2, items.getMaxRows() - 1, 1).setNumberFormat('yyyy/mm/dd hh:mm');

  const menu = ensureSheet_(ss, SHEET.MENU, MENU_HEADERS);
  if (menu.getLastRow() < 2) menu.getRange(2, 1, MENU_SEED.length, MENU_HEADERS.length).setValues(MENU_SEED);
  menu.getRange(2, 7, menu.getMaxRows() - 1, 1)
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Y', 'N'], true).build());

  ensureSheet_(ss, SHEET.LOG, ['時間', '位置', '訊息', '內容']);

  // 觸發器：狀態變更通知、每日日報（先清掉舊的避免重複）
  ScriptApp.getProjectTriggers()
    .filter((t) => ['onStatusEdit', 'sendDailyReport'].indexOf(t.getHandlerFunction()) >= 0)
    .forEach((t) => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('onStatusEdit').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('sendDailyReport').timeBased().everyDays(1).atHour(CONFIG.DAILY_REPORT_HOUR)
    .inTimezone(CONFIG.TIMEZONE).create();

  const def = ss.getSheetByName('工作表1') || ss.getSheetByName('Sheet1');
  if (def && def.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(def);
  ss.setActiveSheet(orders);
  return '設定完成：已建立工作表與觸發器';
}

function ensureSheet_(ss, name, headers) {
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#4e3220').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  return sh;
}

/* =========================== 網站 API =========================== */
function doPost(e) {
  let body = {};
  try {
    body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.action && body.action !== 'createOrder') return json_({ ok: false, error: '不支援的操作' });
    return json_(createOrder_(body));
  } catch (err) {
    if (err && err.userMessage) return json_({ ok: false, error: err.userMessage });
    logError_('doPost', err, body);
    return json_({ ok: false, error: '系統忙碌中，請稍後再試或來電訂購' });
  }
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action === 'status') return json_(lookupStatus_(p.orderNo, p.phone));
  return json_({ ok: true, service: CONFIG.SHOP_NAME + ' 訂單 API', time: now_() });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** 給客人看的錯誤（會原樣顯示在網站上） */
function userError_(msg) { const e = new Error(msg); e.userMessage = msg; return e; }

/* =========================== 建立訂單 =========================== */
function createOrder_(b) {
  // 機器人陷阱：隱藏欄位有值就當作成功但不寫入
  if (b.website) return { ok: true, orderNo: 'JSJ-' + Date.now() };

  // 重送保護：同一個 clientId 只建立一次訂單
  const cache = CacheService.getScriptCache();
  if (b.clientId) {
    const prev = cache.get('cid:' + b.clientId);
    if (prev) return JSON.parse(prev);
  }

  const o = validate_(b);

  const phoneKey = 'ph:' + o.phone;
  const recent = Number(cache.get(phoneKey) || 0);
  if (recent >= CONFIG.MAX_ORDERS_PER_PHONE) throw userError_('下單次數過於頻繁，請稍後再試或直接來電');

  const priced = price_(o);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) throw userError_('目前訂單較多，請稍後再試');
  let orderNo, created;
  try {
    created = new Date();
    orderNo = nextOrderNo_(created);
    const ss = ss_();
    const summary = priced.lines.map((l) => l.name + '×' + l.qty + l.unit).join('、');
    ss.getSheetByName(SHEET.ORDERS).appendRow([
      "'" + orderNo, created, '新訂單', MODE_TEXT[o.mode], o.mode === 'delivery' ? '—' : CONFIG.STORES[o.store],
      o.mode === 'delivery' ? '—' : o.time, safe_(o.name), "'" + o.phone, safe_(o.email), safe_(o.address),
      safe_(o.note), summary, priced.subtotal, priced.fee, priced.total, ''
    ]);
    const itemRows = priced.lines.map((l) => ["'" + orderNo, created, l.id, l.name, l.price, l.qty, l.unit, l.amount]);
    const itemSheet = ss.getSheetByName(SHEET.ITEMS);
    itemSheet.getRange(itemSheet.getLastRow() + 1, 1, itemRows.length, ITEM_HEADERS.length).setValues(itemRows);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  const result = { ok: true, orderNo: orderNo, subtotal: priced.subtotal, fee: priced.fee, total: priced.total,
    lines: priced.lines.map((l) => ({ name: l.name, qty: l.qty, unit: l.unit, amount: l.amount })) };
  if (b.clientId) cache.put('cid:' + b.clientId, JSON.stringify(result), 21600);
  cache.put(phoneKey, String(recent + 1), 600);

  // 寄信失敗不影響訂單成立
  const order = Object.assign({ orderNo: orderNo, created: created }, o, priced);
  try { notifyShop_(order); } catch (err) { logError_('notifyShop', err, orderNo); }
  if (o.email) { try { mailCustomer_(order); } catch (err) { logError_('mailCustomer', err, orderNo); } }
  return result;
}

function validate_(b) {
  const s = (v, max) => String(v == null ? '' : v).trim().slice(0, max);
  const o = {
    mode: s(b.mode, 10), store: s(b.store, 20), time: s(b.time, 20),
    name: s(b.name, 40), phone: s(b.phone, 20).replace(/[\s-]/g, ''), email: s(b.email, 120),
    address: s(b.address, 200), note: s(b.note, 300), items: Array.isArray(b.items) ? b.items.slice(0, 50) : []
  };
  if (!MODE_TEXT[o.mode]) throw userError_('請選擇取餐方式');
  if (!o.name) throw userError_('請填寫姓名');
  if (!/^09\d{8}$/.test(o.phone)) throw userError_('手機號碼格式不正確');
  if (o.email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(o.email)) throw userError_('Email 格式不正確');
  if (o.mode === 'delivery' && !o.address) throw userError_('請填寫宅配地址');
  if (o.mode !== 'delivery' && !CONFIG.STORES[o.store]) throw userError_('請選擇取餐門市');
  if (!o.items.length) throw userError_('購物車是空的');
  return o;
}

/** 依「菜單」工作表計價：不信任網頁送來的價格 */
function price_(o) {
  const rows = ss_().getSheetByName(SHEET.MENU).getDataRange().getValues().slice(1);
  const menu = {};
  rows.forEach((r) => { if (r[0]) menu[String(r[0]).trim()] = { name: r[1], price: Number(r[3]), unit: r[4], min: Number(r[5]) || 1, on: String(r[6]).trim().toUpperCase() !== 'N' }; });

  const merged = {};
  o.items.forEach((it) => {
    const id = String(it && it.id || '').trim(), qty = Math.floor(Number(it && it.qty));
    if (!id) return;
    merged[id] = (merged[id] || 0) + (qty > 0 ? qty : 0);
  });

  const lines = Object.keys(merged).map((id) => {
    const m = menu[id], qty = merged[id];
    if (!m) throw userError_('品項已不存在，請重新整理頁面後再下單');
    if (!m.on) throw userError_('「' + m.name + '」目前暫停供應，請移除後再下單');
    if (qty <= 0 || qty > 2000 || qty % m.min !== 0) throw userError_('「' + m.name + '」數量需為 ' + m.min + ' 的倍數');
    return { id: id, name: m.name, price: m.price, qty: qty, unit: m.unit, amount: m.price * qty };
  });
  if (!lines.length) throw userError_('購物車是空的');

  const subtotal = lines.reduce((t, l) => t + l.amount, 0);
  let fee = 0;
  if (o.mode === 'delivery') fee = subtotal >= CONFIG.FREE_SHIP ? 0 : CONFIG.DELIVERY_FEE;
  else if (o.mode === 'dine') fee = Math.round(subtotal * CONFIG.DINE_RATE);
  return { lines: lines, subtotal: subtotal, fee: fee, total: subtotal + fee };
}

/** 訂單編號：JSJ + 年月日 + 當日流水號，例如 JSJ261116-007（需在鎖內呼叫） */
function nextOrderNo_(d) {
  const day = Utilities.formatDate(d, CONFIG.TIMEZONE, 'yyMMdd');
  const props = PropertiesService.getScriptProperties();
  const key = 'seq-' + day;
  const n = Number(props.getProperty(key) || 0) + 1;
  props.setProperty(key, String(n));
  return 'JSJ' + day + '-' + ('00' + n).slice(-3);
}

/** 防止試算表公式注入（以 = + - @ 開頭的文字） */
function safe_(v) { return /^[=+\-@]/.test(v) ? "'" + v : v; }

/* =========================== 查詢訂單 =========================== */
function lookupStatus_(orderNo, phone) {
  orderNo = String(orderNo || '').trim(); phone = String(phone || '').replace(/[\s-]/g, '');
  if (!orderNo || !phone) return { ok: false, error: '請提供訂單編號與電話' };
  const sh = ss_().getSheetByName(SHEET.ORDERS);
  const hit = sh.getRange(1, COL['訂單編號'], sh.getLastRow(), 1).createTextFinder(orderNo).matchEntireCell(true).findNext();
  if (!hit || hit.getRow() === 1) return { ok: false, error: '查無此訂單' };
  const r = sh.getRange(hit.getRow(), 1, 1, ORDER_HEADERS.length).getValues()[0];
  if (String(r[COL['電話'] - 1]).replace(/^'/, '') !== phone) return { ok: false, error: '查無此訂單' };
  return { ok: true, orderNo: orderNo, status: r[COL['狀態'] - 1], mode: r[COL['取餐方式'] - 1], total: r[COL['總額'] - 1] };
}

/* =========================== 狀態變更通知 =========================== */
/** 安裝式 onEdit 觸發器：在「訂單」改狀態時寄信給客人 */
function onStatusEdit(e) {
  const range = e && e.range;
  if (!range || range.getSheet().getName() !== SHEET.ORDERS || range.getColumn() !== COL['狀態'] || range.getRow() < 2) return;
  const sh = range.getSheet();
  for (let r = range.getRow(); r < range.getRow() + range.getNumRows(); r++) {
    const row = sh.getRange(r, 1, 1, ORDER_HEADERS.length).getValues()[0];
    const status = row[COL['狀態'] - 1], email = row[COL['Email'] - 1];
    const msg = statusMessage_(status, row);
    if (!msg || !email) continue;
    try {
      MailApp.sendEmail({ to: email, subject: '【' + CONFIG.SHOP_NAME + '】訂單 ' + row[0] + ' ' + status, htmlBody: mailFrame_(msg), name: CONFIG.SHOP_NAME });
      const log = sh.getRange(r, COL['通知紀錄']);
      log.setValue((log.getValue() ? log.getValue() + '\n' : '') + now_() + ' 已通知「' + status + '」');
    } catch (err) { logError_('onStatusEdit', err, row[0]); }
  }
}

function statusMessage_(status, row) {
  const name = esc_(row[COL['姓名'] - 1]), no = esc_(row[0]);
  if (status === '可取餐') return '<p>' + name + ' 您好，</p><p>您的訂單 <b>' + no + '</b> 已經煎好囉！請至「' + CONFIG.SHOP_NAME + ' ' + esc_(row[COL['門市'] - 1]) + '」報上訂單編號取餐，趁熱最好吃。</p>';
  if (status === '已出貨') return '<p>' + name + ' 您好，</p><p>您的冷凍宅配訂單 <b>' + no + '</b> 已經出貨，預計 1–2 個工作天送達：' + esc_(row[COL['地址'] - 1]) + '。收到後請立即冷凍保存。</p>';
  if (status === '已取消') return '<p>' + name + ' 您好，</p><p>您的訂單 <b>' + no + '</b> 已取消。如有疑問，歡迎來電與我們聯繫。</p>';
  return '';
}

/* =========================== Email =========================== */
function shopEmail_() { return CONFIG.NOTIFY_EMAIL || Session.getEffectiveUser().getEmail(); }

function orderTable_(o) {
  const rows = o.lines.map((l) => '<tr><td style="padding:4px 8px">' + esc_(l.name) + '</td><td style="padding:4px 8px;text-align:right">' + l.qty + esc_(l.unit) + '</td><td style="padding:4px 8px;text-align:right">' + money_(l.amount) + '</td></tr>').join('');
  const feeLabel = o.mode === 'delivery' ? '運費' : o.mode === 'dine' ? '內用服務費' : '服務費';
  return '<table style="border-collapse:collapse;width:100%;max-width:480px">' + rows +
    '<tr><td colspan="2" style="padding:4px 8px;border-top:1px solid #ddd">小計</td><td style="padding:4px 8px;text-align:right;border-top:1px solid #ddd">' + money_(o.subtotal) + '</td></tr>' +
    '<tr><td colspan="2" style="padding:4px 8px">' + feeLabel + '</td><td style="padding:4px 8px;text-align:right">' + (o.fee ? money_(o.fee) : '免費') + '</td></tr>' +
    '<tr><td colspan="2" style="padding:4px 8px"><b>總額</b></td><td style="padding:4px 8px;text-align:right;color:#e46f2a"><b>' + money_(o.total) + '</b></td></tr></table>';
}

function pickupText_(o) {
  return o.mode === 'delivery' ? '冷凍宅配至：' + esc_(o.address)
    : MODE_TEXT[o.mode] + '｜' + CONFIG.STORES[o.store] + '｜' + esc_(o.time);
}

function notifyShop_(o) {
  const html = '<p><b>新訂單 ' + o.orderNo + '</b>（' + Utilities.formatDate(o.created, CONFIG.TIMEZONE, 'MM/dd HH:mm') + '）</p>' +
    '<p>' + pickupText_(o) + '<br>客人：' + esc_(o.name) + '　' + o.phone + (o.email ? '　' + esc_(o.email) : '') + '</p>' +
    (o.note ? '<p>備註：' + esc_(o.note) + '</p>' : '') + orderTable_(o) +
    '<p><a href="' + ss_().getUrl() + '">開啟訂單試算表</a></p>';
  MailApp.sendEmail({ to: shopEmail_(), subject: '🥟 新訂單 ' + o.orderNo + '｜' + MODE_TEXT[o.mode] + '｜' + money_(o.total), htmlBody: mailFrame_(html), name: CONFIG.SHOP_NAME + ' 訂單系統' });
}

function mailCustomer_(o) {
  const tail = o.mode === 'delivery' ? '我們將於 1–2 個工作天內出貨，出貨時會再寄信通知您。' : '餐點煎好時會再寄信通知您，取餐時請報上訂單編號，現場付款。';
  const html = '<p>' + esc_(o.name) + ' 您好，感謝您的訂購！</p><p>訂單編號：<b>' + o.orderNo + '</b><br>' + pickupText_(o) + '</p>' +
    orderTable_(o) + '<p>' + tail + '</p>';
  MailApp.sendEmail({ to: o.email, subject: '【' + CONFIG.SHOP_NAME + '】訂單確認 ' + o.orderNo, htmlBody: mailFrame_(html), name: CONFIG.SHOP_NAME });
}

function mailFrame_(inner) {
  return '<div style="font-family:\'Noto Sans TC\',sans-serif;color:#3b2618;line-height:1.7;max-width:560px">' +
    '<div style="background:#e46f2a;color:#fff;padding:12px 16px;border-radius:12px 12px 0 0;font-weight:bold;font-size:18px">' + CONFIG.SHOP_NAME + ' 鍋貼專賣</div>' +
    '<div style="border:1px solid #efe3d1;border-top:0;padding:16px;border-radius:0 0 12px 12px">' + inner + '</div></div>';
}

/* =========================== 營業日報 =========================== */
function sendDailyReport() {
  const ss = ss_();
  const today = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy/MM/dd');
  const isToday = (d) => d instanceof Date && Utilities.formatDate(d, CONFIG.TIMEZONE, 'yyyy/MM/dd') === today;

  const orders = ss.getSheetByName(SHEET.ORDERS).getDataRange().getValues().slice(1)
    .filter((r) => isToday(r[COL['建立時間'] - 1]) && r[COL['狀態'] - 1] !== '已取消');
  const cancelled = ss.getSheetByName(SHEET.ORDERS).getDataRange().getValues().slice(1)
    .filter((r) => isToday(r[COL['建立時間'] - 1]) && r[COL['狀態'] - 1] === '已取消').length;
  const validNos = {}; orders.forEach((r) => { validNos[String(r[0])] = true; });

  const revenue = orders.reduce((t, r) => t + Number(r[COL['總額'] - 1] || 0), 0);
  const byMode = {}, byStore = {};
  orders.forEach((r) => {
    const m = r[COL['取餐方式'] - 1], s = r[COL['門市'] - 1], v = Number(r[COL['總額'] - 1] || 0);
    byMode[m] = (byMode[m] || 0) + v; if (s && s !== '—') byStore[s] = (byStore[s] || 0) + v;
  });
  const qty = {};
  ss.getSheetByName(SHEET.ITEMS).getDataRange().getValues().slice(1)
    .filter((r) => validNos[String(r[0])])
    .forEach((r) => { const k = r[3] + '（' + r[6] + '）'; qty[k] = (qty[k] || 0) + Number(r[5] || 0); });
  const top = Object.keys(qty).sort((a, b) => qty[b] - qty[a]).slice(0, 5);

  const list = (obj) => Object.keys(obj).map((k) => '<li>' + esc_(k) + '：' + money_(obj[k]) + '</li>').join('') || '<li>—</li>';
  const html = '<p><b>' + today + ' 營業日報</b></p>' +
    '<p>訂單數：<b>' + orders.length + '</b> 筆（另取消 ' + cancelled + ' 筆）<br>營業額：<b style="color:#e46f2a">' + money_(revenue) + '</b>' +
    (orders.length ? '<br>平均客單價：' + money_(Math.round(revenue / orders.length)) : '') + '</p>' +
    '<p>依取餐方式</p><ul>' + list(byMode) + '</ul><p>依門市</p><ul>' + list(byStore) + '</ul>' +
    '<p>熱銷前 5 名</p><ol>' + (top.map((k) => '<li>' + esc_(k) + ' × ' + qty[k] + '</li>').join('') || '<li>—</li>') + '</ol>';
  MailApp.sendEmail({ to: shopEmail_(), subject: '📊 ' + CONFIG.SHOP_NAME + ' ' + today + ' 營業日報｜' + money_(revenue), htmlBody: mailFrame_(html), name: CONFIG.SHOP_NAME + ' 訂單系統' });
}

/* =========================== 小工具 =========================== */
/** 取得訂單試算表：優先用 SPREADSHEET_ID，否則用綁定的試算表 */
function ss_() {
  if (CONFIG.SPREADSHEET_ID) return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID.trim());
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('找不到試算表：請在 CONFIG.SPREADSHEET_ID 填入試算表 ID，或從試算表的「擴充功能 → Apps Script」建立程式');
  return ss;
}

function now_() { return Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy/MM/dd HH:mm:ss'); }
function money_(n) { return 'NT$' + Number(n).toLocaleString('en-US'); }
function esc_(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }

function logError_(where, err, data) {
  try {
    ss_().getSheetByName(SHEET.LOG)
      .appendRow([new Date(), where, String(err && err.stack || err), JSON.stringify(data || '').slice(0, 2000)]);
  } catch (e) { console.error(where, err); }
}

/** 在編輯器手動執行，測試寫入與寄信（會建立一筆真的測試訂單，測完可刪除該列） */
function testCreateOrder() {
  const res = createOrder_({ mode: 'takeout', store: 'flagship', time: '今天 12:00', name: '測試客人', phone: '0912345678',
    email: shopEmail_(), note: '這是測試訂單', items: [{ id: 'p1', qty: 20 }, { id: 's3', qty: 1 }], clientId: 'test-' + Date.now() });
  console.log(res);
}
