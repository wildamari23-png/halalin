const APP_NAME = 'Halalin Printing';
const APP_COLORS = { primary: '#0B3D91', secondary: '#FFFFFF' };

function doGet() {
  bootstrapSheets_();
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle(APP_NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function bootstrapSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const schemas = [
    { name: 'Users', headers: ['id', 'name', 'username', 'password', 'role', 'active', 'createdAt'] },
    { name: 'Products', headers: ['id', 'name', 'category', 'unit', 'basePrice', 'discountPct', 'imageUrl', 'active'] },
    { name: 'Orders', headers: ['id', 'orderNo', 'userId', 'customerName', 'customerPhone', 'itemsJson', 'subtotal', 'discountTotal', 'grandTotal', 'status', 'createdAt'] },
    { name: 'PriceRules', headers: ['id', 'category', 'minQty', 'unitPrice', 'discountPct', 'active'] },
    { name: 'Settings', headers: ['key', 'value'] }
  ];

  schemas.forEach(schema => {
    let sh = ss.getSheetByName(schema.name);
    if (!sh) sh = ss.insertSheet(schema.name);
    if (sh.getLastRow() === 0) sh.appendRow(schema.headers);
    if (sh.getLastRow() > 0 && sh.getRange(1, 1, 1, schema.headers.length).getValues()[0].join('|') !== schema.headers.join('|')) {
      sh.clear();
      sh.appendRow(schema.headers);
    }
  });

  seedDefaults_(ss);
}

function seedDefaults_(ss) {
  const users = ss.getSheetByName('Users');
  if (users.getLastRow() === 1) {
    users.appendRow([uid_(), 'Super Admin', 'grandadmin', 'admin123', 'grand_admin', true, new Date()]);
    users.appendRow([uid_(), 'Admin Operasional', 'admin', 'admin123', 'admin', true, new Date()]);
  }

  const products = ss.getSheetByName('Products');
  if (products.getLastRow() === 1) {
    [
      ['Undangan Premium', 'Undangan Cetak', 'pcs', 5000, 5],
      ['Banner Vinyl', 'Banner', 'm2', 45000, 10],
      ['Stempel Flash', 'Stempel', 'pcs', 75000, 0],
      ['Undangan Digital Motion', 'Undangan Digital', 'pcs', 150000, 15]
    ].forEach(p => products.appendRow([uid_(), p[0], p[1], p[2], p[3], p[4], '', true]));
  }

  const settings = ss.getSheetByName('Settings');
  if (settings.getLastRow() === 1) {
    settings.appendRow(['brandName', APP_NAME]);
    settings.appendRow(['themePrimary', APP_COLORS.primary]);
    settings.appendRow(['themeSecondary', APP_COLORS.secondary]);
  }
}

function login(username, password) {
  const rows = getSheetData_('Users');
  const user = rows.find(u => u.username === username && u.password === password && String(u.active) === 'true');
  if (!user) throw new Error('Username / password salah.');
  return { id: user.id, name: user.name, role: user.role };
}

function getInitialData() {
  return {
    products: getSheetData_('Products').filter(p => String(p.active) === 'true'),
    orders: getSheetData_('Orders'),
    priceRules: getSheetData_('PriceRules').filter(r => String(r.active) === 'true')
  };
}

function saveOrder(payload) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Orders');
  const id = uid_();
  const orderNo = 'HLN-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
  sh.appendRow([
    id, orderNo, payload.userId, payload.customerName, payload.customerPhone,
    JSON.stringify(payload.items), payload.subtotal, payload.discountTotal,
    payload.grandTotal, 'NEW', new Date()
  ]);

  const invoice = createInvoicePdf_({ ...payload, orderNo });
  return { ok: true, id, orderNo, invoicePdfUrl: invoice.pdfUrl, invoiceImageUrl: invoice.imageUrl };
}

function upsertProduct(product) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Products');
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const idx = data.findIndex((r, i) => i > 0 && r[0] === product.id);
  const row = [product.id || uid_(), product.name, product.category, product.unit, Number(product.basePrice), Number(product.discountPct), product.imageUrl || '', true];
  if (idx > -1) sh.getRange(idx + 1, 1, 1, headers.length).setValues([row]);
  else sh.appendRow(row);
  return { ok: true };
}

function upsertPriceRule(rule) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PriceRules');
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const idx = data.findIndex((r, i) => i > 0 && r[0] === rule.id);
  const row = [rule.id || uid_(), rule.category, Number(rule.minQty), Number(rule.unitPrice), Number(rule.discountPct), true];
  if (idx > -1) sh.getRange(idx + 1, 1, 1, headers.length).setValues([row]);
  else sh.appendRow(row);
  return { ok: true };
}

function generateReports(period) {
  const orders = getSheetData_('Orders');
  const now = new Date();
  const filtered = orders.filter(o => {
    const d = new Date(o.createdAt);
    if (period === 'daily') return sameDay_(d, now);
    if (period === 'monthly') return d.getMonth() === now.getMonth() && d.getFullYear() === now.getFullYear();
    return d.getFullYear() === now.getFullYear();
  });

  const total = filtered.reduce((a, b) => a + Number(b.grandTotal || 0), 0);
  const count = filtered.length;

  const doc = DocumentApp.create(`${APP_NAME} Report ${period} ${new Date().toISOString()}`);
  const body = doc.getBody();
  body.appendParagraph(`${APP_NAME} - Laporan ${period.toUpperCase()}`).setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph(`Jumlah Order: ${count}`);
  body.appendParagraph(`Total Omzet: Rp ${formatNumber_(total)}`);
  filtered.slice(0, 50).forEach(o => body.appendListItem(`${o.orderNo} | ${o.customerName} | Rp ${formatNumber_(Number(o.grandTotal || 0))}`));
  doc.saveAndClose();

  const pdf = DriveApp.getFileById(doc.getId()).getAs(MimeType.PDF);
  const file = DriveApp.createFile(pdf).setName(`Report-${period}-${Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMddHHmm')}.pdf`);
  return { count, total, pdfUrl: file.getUrl() };
}

function createInvoicePdf_(order) {
  const html = `
    <div style="font-family:Arial;padding:16px;">
      <h2 style="color:${APP_COLORS.primary};">Invoice ${order.orderNo}</h2>
      <p>Pelanggan: ${order.customerName} (${order.customerPhone})</p>
      <p>Total: Rp ${formatNumber_(order.grandTotal)}</p>
      <hr/><p>${APP_NAME}</p>
    </div>`;
  const blob = Utilities.newBlob(html, 'text/html', `${order.orderNo}.html`);
  const image = DriveApp.createFile(blob).setName(`${order.orderNo}-invoice.html`);

  const doc = DocumentApp.create(`Invoice-${order.orderNo}`);
  doc.getBody().appendParagraph(`Invoice ${order.orderNo}`).setHeading(DocumentApp.ParagraphHeading.HEADING1);
  doc.getBody().appendParagraph(`Pelanggan: ${order.customerName}`);
  doc.getBody().appendParagraph(`Total: Rp ${formatNumber_(order.grandTotal)}`);
  doc.saveAndClose();
  const pdf = DriveApp.createFile(DriveApp.getFileById(doc.getId()).getAs(MimeType.PDF)).setName(`Invoice-${order.orderNo}.pdf`);

  return { pdfUrl: pdf.getUrl(), imageUrl: image.getUrl() };
}

function getSheetData_(sheetName) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const values = sh.getDataRange().getValues();
  const headers = values.shift();
  return values.map(r => headers.reduce((o, h, i) => (o[h] = r[i], o), {}));
}

function uid_() { return Utilities.getUuid(); }
function formatNumber_(n) { return Number(n || 0).toLocaleString('id-ID'); }
function sameDay_(a, b) { return a.toDateString() === b.toDateString(); }
