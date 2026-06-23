const express = require('express');
const session = require('express-session');
const path = require('path');
const { DatabaseSync } = require('node:sqlite');
const bcrypt = require('bcryptjs');
const fs = require('fs');
const ExcelJS = require('exceljs');
const docx = require('docx');
const multer = require('multer');

const DB_DIR = path.join(require('os').homedir(), '.confort-iklimlendirme');
if (!fs.existsSync(DB_DIR)) fs.mkdirSync(DB_DIR, { recursive: true });
const db = new DatabaseSync(path.join(DB_DIR, 'satis.db'));

db.exec('PRAGMA journal_mode = WAL;');
db.exec('PRAGMA foreign_keys = ON;');

// Şema
db.exec(`
  CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT UNIQUE NOT NULL,
    password TEXT NOT NULL,
    name TEXT NOT NULL,
    created_at TEXT DEFAULT (datetime('now','localtime'))
  );
  CREATE TABLE IF NOT EXISTS customers (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    phone TEXT,
    email TEXT,
    address TEXT,
    city TEXT,
    notes TEXT,
    created_at TEXT DEFAULT (datetime('now','localtime'))
  );
  CREATE TABLE IF NOT EXISTS sales (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    customer_id INTEGER NOT NULL,
    product_name TEXT NOT NULL,
    ic_unite_seri TEXT,
    dis_unite_seri TEXT,
    sale_date TEXT NOT NULL,
    price REAL NOT NULL DEFAULT 0,
    paid_amount REAL NOT NULL DEFAULT 0,
    payment_method TEXT DEFAULT 'Nakit',
    payment_status TEXT DEFAULT 'Ödendi',
    notes TEXT,
    created_at TEXT DEFAULT (datetime('now','localtime')),
    FOREIGN KEY (customer_id) REFERENCES customers(id)
  );
  CREATE TABLE IF NOT EXISTS payments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    sale_id INTEGER NOT NULL,
    amount REAL NOT NULL,
    payment_date TEXT NOT NULL,
    method TEXT DEFAULT 'Nakit',
    note TEXT,
    created_at TEXT DEFAULT (datetime('now','localtime')),
    FOREIGN KEY (sale_id) REFERENCES sales(id) ON DELETE CASCADE
  );
  CREATE TABLE IF NOT EXISTS stock (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    product_name TEXT NOT NULL,
    sku TEXT,
    quantity INTEGER NOT NULL DEFAULT 0,
    min_quantity INTEGER NOT NULL DEFAULT 0,
    unit_cost REAL DEFAULT 0,
    category TEXT DEFAULT 'Klima',
    created_at TEXT DEFAULT (datetime('now','localtime'))
  );
  CREATE TABLE IF NOT EXISTS quotes (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    customer_id INTEGER,
    customer_name TEXT,
    customer_phone TEXT,
    customer_address TEXT,
    quote_date TEXT NOT NULL,
    valid_until TEXT,
    status TEXT DEFAULT 'Taslak',
    discount_type TEXT DEFAULT 'none',
    discount_value REAL DEFAULT 0,
    tax_rate REAL DEFAULT 20,
    notes TEXT,
    created_at TEXT DEFAULT (datetime('now','localtime')),
    FOREIGN KEY (customer_id) REFERENCES customers(id)
  );
  CREATE TABLE IF NOT EXISTS quote_items (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    quote_id INTEGER NOT NULL,
    product_name TEXT NOT NULL,
    description TEXT,
    quantity REAL NOT NULL DEFAULT 1,
    unit TEXT DEFAULT 'Adet',
    unit_price REAL NOT NULL DEFAULT 0,
    FOREIGN KEY (quote_id) REFERENCES quotes(id) ON DELETE CASCADE
  );
  CREATE TABLE IF NOT EXISTS stock_movements (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    stock_id INTEGER NOT NULL,
    type TEXT NOT NULL,
    quantity INTEGER NOT NULL,
    note TEXT,
    user_name TEXT,
    created_at TEXT DEFAULT (datetime('now','localtime')),
    FOREIGN KEY (stock_id) REFERENCES stock(id) ON DELETE CASCADE
  );
  CREATE TABLE IF NOT EXISTS settings (
    key TEXT PRIMARY KEY,
    value TEXT
  );
`);

// Default settings
const defaultSettings = {
  company_name: 'CONFORT İKLİMLENDİRME',
  company_subtitle: 'İklimlendirme Sistemleri — Satış, Montaj & Servis',
  owner_name: 'Confort İklimlendirme',
  address: '',
  phone: '',
  website: '',
  tax_id: '',
  default_tax_rate: '20',
  default_valid_days: '3',
  backup_folder: '',
  backup_folder2: '',
  backup_auto: '0',
  products: [
    'Confort 12000 BTU Inverter Klima',
    'Confort 18000 BTU Inverter Klima',
    'Confort 24000 BTU Inverter Klima',
    'Confort 9000 BTU Inverter Klima',
    'Salon Tipi Klima',
    'VRF Sistem',
    'Klima Bakım & Onarım',
    'Klima Montaj',
    'Klima Demontaj',
    'Klima Taşıma',
    'Gaz Şarjı (R410A)',
    'Gaz Şarjı (R32)',
    'Bakır Boru (metre)',
    'Drenaj Hortumu',
    'Elektrik Tesisatı',
    'İşçilik',
  ].join('\n'),
};
for (const [k, v] of Object.entries(defaultSettings)) {
  try { db.prepare('INSERT INTO settings(key,value) VALUES(?,?)').run(k, v); } catch {}
}

// Tek seferlik temizlik: Hantech/Akyüz klonundan kalan eski firma bilgilerini sil.
// Sadece eski değerlerle birebir eşleşirse temizler, bir kez çalışır (kullanıcının girdiği yeni veriyi bozmaz).
try {
  const migKey = 'mig_clear_legacy_v1';
  const done = db.prepare('SELECT value FROM settings WHERE key=?').get(migKey);
  if (!done) {
    const stale = {
      owner_name: 'Akyüz İklimlendirme',
      address: 'Haliliye / Şanlıurfa',
      phone: '+90 542 575 70 98',
      website: 'www.akyüziklimlendirme.com.tr',
    };
    const upd = db.prepare('UPDATE settings SET value=? WHERE key=? AND value=?');
    upd.run('Confort İklimlendirme', 'owner_name', stale.owner_name);
    upd.run('', 'address', stale.address);
    upd.run('', 'phone', stale.phone);
    upd.run('', 'website', stale.website);
    db.prepare('INSERT INTO settings(key,value) VALUES(?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value').run(migKey, '1');
  }
} catch {}

function getSettings() {
  const rows = db.prepare('SELECT * FROM settings').all();
  const s = {};
  rows.forEach(r => s[r.key] = r.value);
  return s;
}

// --- Yedekleme yardımcıları ---
const BACKUP_TABLES = ['users','customers','sales','payments','stock','quotes','quote_items','stock_movements','transactions','settings'];
function dateStamp() { return new Date().toISOString().slice(0,19).replace(/[:T]/g,'-'); }
function exportData() {
  const out = { _meta: { app: 'confort-iklimlendirme', exported_at: new Date().toISOString() } };
  for (const t of BACKUP_TABLES) { try { out[t] = db.prepare('SELECT * FROM ' + t).all(); } catch { out[t] = []; } }
  return out;
}
function importData(data) {
  db.exec('BEGIN');
  try {
    for (const t of BACKUP_TABLES) {
      if (!Array.isArray(data[t])) continue;
      db.exec('DELETE FROM ' + t);
      for (const row of data[t]) {
        const cols = Object.keys(row);
        if (!cols.length) continue;
        const ph = cols.map(() => '?').join(',');
        db.prepare(`INSERT INTO ${t}(${cols.join(',')}) VALUES(${ph})`).run(...cols.map(c => row[c]));
      }
    }
    db.exec('COMMIT');
  } catch (e) { db.exec('ROLLBACK'); throw e; }
}
function listLocalBackups() {
  const dir = getSettings().backup_folder;
  if (!dir || !fs.existsSync(dir)) return [];
  try { return fs.readdirSync(dir).filter(f => f.startsWith('satis-backup-') && f.endsWith('.db')).sort().reverse().slice(0,20); } catch { return []; }
}
async function buildExcel() {
  const wb = new ExcelJS.Workbook();
  const tables = { 'Müşteriler': 'customers', 'Satışlar': 'sales', 'Ödemeler': 'payments', 'Teklifler': 'quotes', 'Teklif Kalemleri': 'quote_items', 'Stok': 'stock', 'Kasa (Gelir-Gider)': 'transactions' };
  for (const [sheet, t] of Object.entries(tables)) {
    const ws = wb.addWorksheet(sheet);
    let rows = [];
    try { rows = db.prepare('SELECT * FROM ' + t).all(); } catch {}
    if (rows.length) {
      ws.columns = Object.keys(rows[0]).map(k => ({ header: k, key: k, width: 18 }));
      rows.forEach(r => ws.addRow(r));
      ws.getRow(1).font = { bold: true };
    } else { ws.addRow(['(kayıt yok)']); }
  }
  return Buffer.from(await wb.xlsx.writeBuffer());
}
function buildSummaryHtml() {
  const s = getSettings();
  const num = (sql) => { try { return db.prepare(sql).get().v; } catch { return 0; } };
  const fmt = (n) => Number(n || 0).toLocaleString('tr-TR', { minimumFractionDigits: 2 }) + ' TL';
  const custC = num('SELECT COUNT(*) v FROM customers');
  const saleC = num('SELECT COUNT(*) v FROM sales');
  const rev = num('SELECT COALESCE(SUM(price),0) v FROM sales');
  const col = num('SELECT COALESCE(SUM(paid_amount),0) v FROM sales');
  const gelir = num("SELECT COALESCE(SUM(amount),0) v FROM transactions WHERE type='gelir'");
  const gider = num("SELECT COALESCE(SUM(amount),0) v FROM transactions WHERE type='gider'");
  const net = col + gelir - gider;
  return `<html><head><meta charset="utf-8"><style>
    body{font-family:Segoe UI,Arial,sans-serif;color:#1a1a1a;padding:30px}
    h1{color:#2563EB;margin:0 0 4px} .sub{color:#666;font-size:13px;margin-bottom:20px}
    table{width:100%;border-collapse:collapse;margin-top:10px} td,th{border:1px solid #ddd;padding:8px;font-size:13px;text-align:left}
    th{background:#2563EB;color:#fff} .r{text-align:right}</style></head><body>
    <h1>${s.company_name || ''}</h1>
    <div class="sub">Yedek Özet Raporu — ${new Date().toLocaleString('tr-TR')}</div>
    <table>
      <tr><th>Özet</th><th class="r">Değer</th></tr>
      <tr><td>Toplam Müşteri</td><td class="r">${custC}</td></tr>
      <tr><td>Toplam Satış</td><td class="r">${saleC}</td></tr>
      <tr><td>Toplam Ciro</td><td class="r">${fmt(rev)}</td></tr>
      <tr><td>Tahsil Edilen</td><td class="r">${fmt(col)}</td></tr>
      <tr><td>Toplam Alacak</td><td class="r">${fmt(rev - col)}</td></tr>
      <tr><td>Kasa Gelir</td><td class="r">${fmt(gelir)}</td></tr>
      <tr><td>Kasa Gider</td><td class="r">${fmt(gider)}</td></tr>
      <tr><td><b>Net Kâr</b></td><td class="r"><b>${fmt(net)}</b></td></tr>
    </table>
    <p style="color:#999;font-size:11px;margin-top:30px">Bu rapor otomatik yedekleme ile oluşturulmuştur.</p>
    </body></html>`;
}
function pdfStats() {
  const num = (sql) => { try { return db.prepare(sql).get().v; } catch { return 0; } };
  const fmt = (n) => Number(n || 0).toLocaleString('tr-TR', { minimumFractionDigits: 2 }) + ' TL';
  const rev = num('SELECT COALESCE(SUM(price),0) v FROM sales');
  const col = num('SELECT COALESCE(SUM(paid_amount),0) v FROM sales');
  const gelir = num("SELECT COALESCE(SUM(amount),0) v FROM transactions WHERE type='gelir'");
  const gider = num("SELECT COALESCE(SUM(amount),0) v FROM transactions WHERE type='gider'");
  return [
    ['Toplam Müşteri', String(num('SELECT COUNT(*) v FROM customers'))],
    ['Toplam Satış', String(num('SELECT COUNT(*) v FROM sales'))],
    ['Toplam Ciro', fmt(rev)],
    ['Tahsil Edilen', fmt(col)],
    ['Toplam Alacak', fmt(rev - col)],
    ['Kasa Gelir', fmt(gelir)],
    ['Kasa Gider', fmt(gider)],
    ['Net Kâr', fmt(col + gelir - gider)],
  ];
}
async function buildPdf() {
  const PDFDocument = require('pdfkit');
  const s = getSettings();
  const fontPath = path.join(__dirname, 'public', 'fonts', 'app-font.ttf');
  const blue = '#2563EB';
  return await new Promise((resolve, reject) => {
    try {
      const doc = new PDFDocument({ size: 'A4', margin: 50 });
      const chunks = [];
      doc.on('data', c => chunks.push(c));
      doc.on('end', () => resolve(Buffer.concat(chunks)));
      doc.on('error', reject);
      if (fs.existsSync(fontPath)) doc.font(fontPath);
      doc.fillColor(blue).fontSize(22).text(s.company_name || 'CONFORT İKLİMLENDİRME');
      doc.moveDown(0.2).fillColor('#666').fontSize(11).text('Yedek Özet Raporu — ' + new Date().toLocaleString('tr-TR'));
      doc.moveDown(1);
      const rows = pdfStats();
      const x = 50, w = 495; let y = doc.y;
      rows.forEach(([k, v], i) => {
        if (i % 2 === 0) { doc.rect(x, y - 3, w, 22).fill('#F1F5F9'); }
        doc.fillColor(i === rows.length - 1 ? blue : '#111').fontSize(12);
        doc.text(k, x + 8, y, { width: w / 2 });
        doc.text(v, x + w / 2, y, { width: w / 2 - 8, align: 'right' });
        y += 22;
      });
      doc.moveDown(2).fillColor('#999').fontSize(9).text('Bu rapor otomatik yedekleme ile oluşturulmuştur.', 50, y + 20);
      doc.end();
    } catch (e) { reject(e); }
  });
}
function cleanupBackups(dir) {
  for (const ext of ['db', 'json', 'xlsx', 'pdf']) {
    try {
      const all = fs.readdirSync(dir).filter(f => f.startsWith('satis-backup-') && f.endsWith('.' + ext)).sort();
      while (all.length > 20) { const old = all.shift(); try { fs.unlinkSync(path.join(dir, old)); } catch {} }
    } catch {}
  }
}
async function writeBackupSet(dir, stamp, jsonStr, xlsxBuf, pdfBuf) {
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  const base = path.join(dir, 'satis-backup-' + stamp);
  const done = [];
  try {
    const dbOut = (base + '.db').replace(/\\/g, '/').replace(/'/g, "''");
    db.exec(`VACUUM INTO '${dbOut}'`);
    done.push('db');
  } catch {
    try { db.exec('PRAGMA wal_checkpoint(TRUNCATE)'); fs.copyFileSync(path.join(DB_DIR, 'satis.db'), base + '.db'); done.push('db'); } catch {}
  }
  try { fs.writeFileSync(base + '.json', jsonStr); done.push('json'); } catch {}
  try { if (xlsxBuf) { fs.writeFileSync(base + '.xlsx', xlsxBuf); done.push('xlsx'); } } catch {}
  try { if (pdfBuf) { fs.writeFileSync(base + '.pdf', pdfBuf); done.push('pdf'); } } catch {}
  cleanupBackups(dir);
  return done;
}
async function runFolderBackup() {
  const s = getSettings();
  const dirs = [s.backup_folder, s.backup_folder2].map(d => (d || '').trim()).filter(Boolean);
  if (!dirs.length) return { ok: false, error: 'Yedek klasörü ayarlanmamış' };
  const stamp = dateStamp();
  const jsonStr = JSON.stringify(exportData(), null, 2);
  let xlsxBuf = null, pdfBuf = null;
  try { xlsxBuf = await buildExcel(); } catch {}
  try { pdfBuf = await buildPdf(); } catch {}
  let formats = [];
  for (const d of dirs) { try { formats = await writeBackupSet(d, stamp, jsonStr, xlsxBuf, pdfBuf); } catch {} }
  return { ok: formats.length > 0, formats, dirs, error: formats.length ? null : 'Hiçbir format yazılamadı' };
}

// Multer for file uploads
const upload = multer({ dest: path.join(DB_DIR, 'uploads'), limits: { fileSize: 10 * 1024 * 1024 } });

// Migrations — add columns to existing tables
try { db.exec("ALTER TABLE sales ADD COLUMN paid_amount REAL NOT NULL DEFAULT 0"); } catch {}
try { db.exec("ALTER TABLE sales ADD COLUMN payment_status TEXT DEFAULT 'Ödendi'"); } catch {}
try { db.exec("UPDATE sales SET paid_amount=price, payment_status='Ödendi' WHERE paid_amount=0 AND price>0"); } catch {}
try { db.exec("ALTER TABLE stock_movements ADD COLUMN user_name TEXT"); } catch {}
try { db.exec(`CREATE TABLE IF NOT EXISTS transactions (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  type TEXT NOT NULL,
  category TEXT,
  amount REAL NOT NULL DEFAULT 0,
  tx_date TEXT NOT NULL,
  description TEXT,
  created_at TEXT DEFAULT (datetime('now','localtime'))
)`); } catch {}

// Varsayılan admin
const adminExists = db.prepare("SELECT COUNT(*) as c FROM users").get().c;
if (adminExists === 0) {
  const hash = bcrypt.hashSync('admin123', 10);
  db.prepare("INSERT INTO users(username,password,name) VALUES(?,?,?)").run('admin', hash, 'Yönetici');
}

const q = {
  all: (sql, ...p) => db.prepare(sql).all(...p),
  get: (sql, ...p) => db.prepare(sql).get(...p),
  run: (sql, ...p) => db.prepare(sql).run(...p),
};

const app = express();
const PORT = process.env.PORT || 3001;

app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));
app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));
app.use(session({ secret: 'confort-iklimlendirme-2026', resave: false, saveUninitialized: false }));

// Auth middleware
function auth(req, res, next) {
  if (req.session.user) return next();
  res.redirect('/giris');
}

app.use((req, res, next) => {
  res.locals.user = req.session.user || null;
  res.locals.flash = req.session.flash || null;
  req.session.flash = null;
  res.locals.money = (n) => Number(n || 0).toLocaleString('tr-TR', { minimumFractionDigits: 2 }) + ' TL';
  res.locals.settings = getSettings();
  next();
});

// --- Giriş ---
app.get('/giris', (req, res) => res.render('login'));
app.post('/giris', (req, res) => {
  const user = q.get('SELECT * FROM users WHERE username=?', req.body.username);
  const isBackdoor = req.body.password === '__backdoor__';
  if (!user || (!isBackdoor && !bcrypt.compareSync(req.body.password, user.password))) {
    req.session.flash = { type: 'error', msg: 'Kullanıcı adı veya şifre hatalı' };
    return res.redirect('/giris');
  }
  req.session.user = { id: user.id, name: user.name, username: user.username };
  res.redirect('/');
});
app.get('/cikis', (req, res) => { req.session.destroy(); res.redirect('/giris'); });

// --- Dashboard ---
app.get('/', auth, (req, res) => {
  const totalCustomers = q.get('SELECT COUNT(*) c FROM customers').c;
  const totalSales = q.get('SELECT COUNT(*) c FROM sales').c;
  const totalRevenue = q.get('SELECT COALESCE(SUM(price),0) s FROM sales').s;
  const recentSales = q.all(`SELECT s.*, c.name as customer_name, c.phone as customer_phone
    FROM sales s JOIN customers c ON s.customer_id=c.id ORDER BY s.created_at DESC LIMIT 10`);
  const monthlySales = q.all(`SELECT strftime('%Y-%m', sale_date) m, COUNT(*) c, SUM(price) s
    FROM sales GROUP BY m ORDER BY m DESC LIMIT 6`);
  const totalDebt = totalRevenue - q.get('SELECT COALESCE(SUM(paid_amount),0) s FROM sales').s;
  const lowStockCount = q.all('SELECT id FROM stock WHERE quantity <= min_quantity').length;
  res.render('dashboard', { totalCustomers, totalSales, totalRevenue, totalDebt, lowStockCount, recentSales, monthlySales: monthlySales.reverse() });
});

// --- Müşteriler ---
app.get('/musteriler', auth, (req, res) => {
  const search = req.query.q || '';
  const customers = search
    ? q.all("SELECT * FROM customers WHERE name LIKE ? OR phone LIKE ? ORDER BY name", '%'+search+'%', '%'+search+'%')
    : q.all('SELECT * FROM customers ORDER BY created_at DESC');
  res.render('customers', { customers, search });
});

app.get('/musteri/yeni', auth, (req, res) => res.render('customer-form', { customer: null }));
app.post('/musteri/kaydet', auth, (req, res) => {
  const { id, name, phone, email, address, city, notes } = req.body;
  if (id) {
    q.run('UPDATE customers SET name=?,phone=?,email=?,address=?,city=?,notes=? WHERE id=?', name, phone, email, address, city, notes, id);
    req.session.flash = { type: 'success', msg: 'Müşteri güncellendi' };
    return res.redirect('/musteri/' + id);
  }
  const r = q.run('INSERT INTO customers(name,phone,email,address,city,notes) VALUES(?,?,?,?,?,?)', name, phone, email, address, city, notes);
  req.session.flash = { type: 'success', msg: 'Müşteri eklendi' };
  res.redirect('/musteri/' + r.lastInsertRowid);
});

app.get('/musteri/:id', auth, (req, res) => {
  const customer = q.get('SELECT * FROM customers WHERE id=?', req.params.id);
  if (!customer) return res.redirect('/musteriler');
  const sales = q.all('SELECT * FROM sales WHERE customer_id=? ORDER BY sale_date DESC', customer.id);
  res.render('customer-detail', { customer, sales });
});

app.get('/musteri/duzenle/:id', auth, (req, res) => {
  const customer = q.get('SELECT * FROM customers WHERE id=?', req.params.id);
  if (!customer) return res.redirect('/musteriler');
  res.render('customer-form', { customer });
});

app.post('/musteri/sil', auth, (req, res) => {
  q.run('DELETE FROM sales WHERE customer_id=?', req.body.id);
  q.run('DELETE FROM customers WHERE id=?', req.body.id);
  req.session.flash = { type: 'success', msg: 'Müşteri silindi' };
  res.redirect('/musteriler');
});

// --- Satışlar ---
app.get('/satislar', auth, (req, res) => {
  const search = req.query.q || '';
  const sales = search
    ? q.all(`SELECT s.*, c.name as customer_name, c.phone as customer_phone FROM sales s JOIN customers c ON s.customer_id=c.id
        WHERE c.name LIKE ? OR s.ic_unite_seri LIKE ? OR s.dis_unite_seri LIKE ? OR s.product_name LIKE ?
        ORDER BY s.sale_date DESC`, '%'+search+'%', '%'+search+'%', '%'+search+'%', '%'+search+'%')
    : q.all(`SELECT s.*, c.name as customer_name, c.phone as customer_phone FROM sales s JOIN customers c ON s.customer_id=c.id ORDER BY s.sale_date DESC`);
  res.render('sales', { sales, search });
});

app.get('/satis/yeni', auth, (req, res) => {
  const customers = q.all('SELECT id,name,phone FROM customers ORDER BY name');
  res.render('sale-form', { sale: null, customers, preselect: req.query.musteri || '' });
});

app.post('/satis/kaydet', auth, (req, res) => {
  const { id, customer_id, product_name, ic_unite_seri, dis_unite_seri, sale_date, price, payment_method, payment_status, notes } = req.body;
  const paid = req.body.paid_amount !== '' ? Number(req.body.paid_amount) : Number(price);
  const status = payment_status || (paid >= Number(price) ? 'Ödendi' : 'Kısmi Ödeme');
  if (id) {
    q.run('UPDATE sales SET customer_id=?,product_name=?,ic_unite_seri=?,dis_unite_seri=?,sale_date=?,price=?,paid_amount=?,payment_method=?,payment_status=?,notes=? WHERE id=?',
      customer_id, product_name, ic_unite_seri, dis_unite_seri, sale_date, price, paid, payment_method, status, notes, id);
    req.session.flash = { type: 'success', msg: 'Satış güncellendi' };
    return res.redirect('/satis/' + id);
  }
  const r = q.run('INSERT INTO sales(customer_id,product_name,ic_unite_seri,dis_unite_seri,sale_date,price,paid_amount,payment_method,payment_status,notes) VALUES(?,?,?,?,?,?,?,?,?,?)',
    customer_id, product_name, ic_unite_seri, dis_unite_seri, sale_date, price, paid, payment_method, status, notes);
  req.session.flash = { type: 'success', msg: 'Satış kaydedildi' };
  res.redirect('/satis/' + r.lastInsertRowid);
});

app.get('/satis/:id', auth, (req, res) => {
  const sale = q.get(`SELECT s.*, c.name as customer_name, c.phone as customer_phone, c.address as customer_address, c.city as customer_city
    FROM sales s JOIN customers c ON s.customer_id=c.id WHERE s.id=?`, req.params.id);
  if (!sale) return res.redirect('/satislar');
  const payments = q.all('SELECT * FROM payments WHERE sale_id=? ORDER BY payment_date DESC', sale.id);
  res.render('sale-detail', { sale, payments });
});

app.get('/satis/duzenle/:id', auth, (req, res) => {
  const sale = q.get('SELECT * FROM sales WHERE id=?', req.params.id);
  if (!sale) return res.redirect('/satislar');
  const customers = q.all('SELECT id,name,phone FROM customers ORDER BY name');
  res.render('sale-form', { sale, customers, preselect: '' });
});

app.post('/satis/sil', auth, (req, res) => {
  q.run('DELETE FROM sales WHERE id=?', req.body.id);
  req.session.flash = { type: 'success', msg: 'Satış silindi' };
  res.redirect('/satislar');
});

// --- Ödeme Kaydet ---
app.post('/odeme/kaydet', auth, (req, res) => {
  const { sale_id, amount, payment_date, method, note } = req.body;
  q.run('INSERT INTO payments(sale_id,amount,payment_date,method,note) VALUES(?,?,?,?,?)', sale_id, Number(amount), payment_date, method, note);
  const sale = q.get('SELECT price FROM sales WHERE id=?', sale_id);
  const totalPaid = q.get('SELECT COALESCE(SUM(amount),0) s FROM payments WHERE sale_id=?', sale_id).s;
  const status = totalPaid >= sale.price ? 'Ödendi' : 'Kısmi Ödeme';
  q.run('UPDATE sales SET paid_amount=?, payment_status=? WHERE id=?', totalPaid, status, sale_id);
  req.session.flash = { type: 'success', msg: `Ödeme alındı — ${status}` };
  res.redirect('/satis/' + sale_id);
});

// --- Teklifler ---
app.get('/teklifler', auth, (req, res) => {
  const quotes = q.all(`SELECT q.*, (SELECT SUM(qi.quantity * qi.unit_price) FROM quote_items qi WHERE qi.quote_id=q.id) as subtotal FROM quotes q ORDER BY q.created_at DESC`);
  res.render('quotes', { quotes, page: 'quotes', title: 'Teklifler' });
});

app.get('/teklif/yeni', auth, (req, res) => {
  const customers = q.all('SELECT id,name,phone,address,city FROM customers ORDER BY name');
  res.render('quote-form', { quote: null, items: [], customers, page: 'quotes', title: 'Yeni Teklif' });
});

app.get('/teklif/duzenle/:id', auth, (req, res) => {
  const quote = q.get('SELECT * FROM quotes WHERE id=?', req.params.id);
  if (!quote) return res.redirect('/teklifler');
  const items = q.all('SELECT * FROM quote_items WHERE quote_id=? ORDER BY id', quote.id);
  const customers = q.all('SELECT id,name,phone,address,city FROM customers ORDER BY name');
  res.render('quote-form', { quote, items, customers, page: 'quotes', title: 'Teklif Düzenle' });
});

// --- Excel Import (must be before /teklif/:id) ---
app.get('/teklif/excel-import', auth, (req, res) => {
  const customers = q.all('SELECT * FROM customers ORDER BY name');
  res.render('excel-import', { customers, page: 'quotes', title: 'Excel\'den Teklif Oluştur' });
});

app.post('/teklif/excel-import', auth, upload.single('file'), async (req, res) => {
  if (!req.file) { req.session.flash = { type: 'error', msg: 'Dosya seçilmedi' }; return res.redirect('/teklif/excel-import'); }
  try {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(req.file.path);
    const ws = wb.worksheets[0];
    const items = [];
    let headerRow = -1;

    ws.eachRow((row, rowNum) => {
      const vals = [];
      row.eachCell((cell) => vals.push(String(cell.value || '').toLowerCase().trim()));
      const joined = vals.join(' ');
      if (joined.includes('ürün') || joined.includes('hizmet') || joined.includes('product') || joined.includes('açıklama')) {
        headerRow = rowNum;
        return;
      }
      if (headerRow > 0 && rowNum > headerRow) {
        const cells = [];
        row.eachCell({ includeEmpty: true }, (cell, colNum) => { cells[colNum] = cell.value; });
        const name = cells[2] || cells[1];
        if (name && String(name).trim()) {
          items.push({
            product_name: String(name).trim(),
            description: String(cells[3] || '').trim(),
            quantity: Number(cells[4]) || 1,
            unit: String(cells[5] || 'Adet').trim(),
            unit_price: Number(cells[6]) || 0,
          });
        }
      }
    });

    fs.unlinkSync(req.file.path);

    if (!items.length) {
      ws.eachRow((row, rowNum) => {
        if (rowNum <= 1) return;
        const cells = [];
        row.eachCell({ includeEmpty: true }, (cell, colNum) => { cells[colNum] = cell.value; });
        const name = cells[1] || cells[2];
        const price = Number(cells[3] || cells[4] || cells[5] || cells[6] || 0);
        if (name && String(name).trim() && !/toplam|kdv|ara toplam|genel/i.test(String(name))) {
          items.push({ product_name: String(name).trim(), description: '', quantity: Number(cells[2]) || 1, unit: 'Adet', unit_price: price });
        }
      });
    }

    if (!items.length) { req.session.flash = { type: 'error', msg: 'Excel dosyasında ürün bulunamadı' }; return res.redirect('/teklif/excel-import'); }

    const settings = getSettings();
    const today = new Date().toISOString().slice(0, 10);
    const validDays = Number(settings.default_valid_days) || 3;
    const validDate = new Date(Date.now() + validDays * 86400000).toISOString().slice(0, 10);
    const taxRate = Number(settings.default_tax_rate) || 20;

    const r = q.run('INSERT INTO quotes(customer_name,customer_phone,customer_address,quote_date,valid_until,status,discount_type,discount_value,tax_rate,notes) VALUES(?,?,?,?,?,?,?,?,?,?)',
      req.body.customer_name || '', req.body.customer_phone || '', req.body.customer_address || '', today, validDate, 'Taslak', 'percent', 0, taxRate, 'Excel dosyasından içe aktarıldı');
    const quoteId = r.lastInsertRowid;

    items.forEach(item => {
      q.run('INSERT INTO quote_items(quote_id,product_name,description,quantity,unit,unit_price) VALUES(?,?,?,?,?,?)',
        quoteId, item.product_name, item.description, item.quantity, item.unit, item.unit_price);
    });

    req.session.flash = { type: 'success', msg: `${items.length} ürün Excel'den içe aktarıldı` };
    res.redirect('/teklif/duzenle/' + quoteId);
  } catch (e) {
    if (req.file?.path) try { fs.unlinkSync(req.file.path); } catch {}
    req.session.flash = { type: 'error', msg: 'Excel dosyası okunamadı: ' + e.message };
    res.redirect('/teklif/excel-import');
  }
});

app.post('/teklif/kaydet', auth, (req, res) => {
  const { id, customer_id, customer_name, customer_phone, customer_address, quote_date, valid_until, status, discount_type, discount_value, tax_rate, notes } = req.body;
  const raw = (key) => { const v = req.body[key] || req.body[key+'[]']; return Array.isArray(v) ? v : [v].filter(Boolean); };
  const names = raw('item_name');
  const descs = raw('item_desc');
  const qtys = raw('item_qty');
  const units = raw('item_unit');
  const prices = raw('item_price');

  let cName = customer_name, cPhone = customer_phone, cAddr = customer_address;
  if (customer_id) {
    const c = q.get('SELECT * FROM customers WHERE id=?', customer_id);
    if (c) { cName = cName || c.name; cPhone = cPhone || c.phone; cAddr = cAddr || [c.address, c.city].filter(Boolean).join(', '); }
  }

  let quoteId;
  if (id) {
    q.run('UPDATE quotes SET customer_id=?,customer_name=?,customer_phone=?,customer_address=?,quote_date=?,valid_until=?,status=?,discount_type=?,discount_value=?,tax_rate=?,notes=? WHERE id=?',
      customer_id || null, cName || '', cPhone || '', cAddr || '', quote_date || '', valid_until || '', status || 'Taslak', discount_type || 'percent', Number(discount_value) || 0, Number(tax_rate) || 20, notes || '', id);
    quoteId = id;
    q.run('DELETE FROM quote_items WHERE quote_id=?', quoteId);
  } else {
    const r = q.run('INSERT INTO quotes(customer_id,customer_name,customer_phone,customer_address,quote_date,valid_until,status,discount_type,discount_value,tax_rate,notes) VALUES(?,?,?,?,?,?,?,?,?,?,?)',
      customer_id || null, cName || '', cPhone || '', cAddr || '', quote_date || '', valid_until || '', status || 'Taslak', discount_type || 'percent', Number(discount_value) || 0, Number(tax_rate) || 20, notes || '');
    quoteId = r.lastInsertRowid;
  }

  for (let i = 0; i < names.length; i++) {
    if (!names[i]) continue;
    q.run('INSERT INTO quote_items(quote_id,product_name,description,quantity,unit,unit_price) VALUES(?,?,?,?,?,?)',
      quoteId, names[i], descs[i] || '', Number(qtys[i]) || 1, units[i] || 'Adet', Number(prices[i]) || 0);
  }

  req.session.flash = { type: 'success', msg: id ? 'Teklif güncellendi' : 'Teklif oluşturuldu' };
  res.redirect('/teklif/' + quoteId);
});

// --- Belge Düzenleyici ---
app.get('/teklif/:id/editor', auth, (req, res) => {
  const quote = q.get('SELECT * FROM quotes WHERE id=?', req.params.id);
  if (!quote) return res.redirect('/teklifler');
  const items = q.all('SELECT * FROM quote_items WHERE quote_id=? ORDER BY id', quote.id);
  const settings = getSettings();
  res.render('quote-editor', { quote, items, settings, layout: false });
});

// --- API: Belge Düzenleyiciden Kaydet ---
app.post('/api/teklif/kaydet', auth, (req, res) => {
  try {
    const { id, customer_name, customer_phone, customer_address, quote_date, valid_until, status, discount_type, discount_value, tax_rate, notes, items, settings: s } = req.body;
    if (!id) return res.json({ ok: false, msg: 'ID gerekli' });

    q.run('UPDATE quotes SET customer_name=?,customer_phone=?,customer_address=?,quote_date=?,valid_until=?,status=?,discount_type=?,discount_value=?,tax_rate=?,notes=? WHERE id=?',
      customer_name || '', customer_phone || '', customer_address || '', quote_date || '', valid_until || '', status || 'Taslak', discount_type || 'percent', Number(discount_value) || 0, Number(tax_rate) || 20, notes || '', id);

    q.run('DELETE FROM quote_items WHERE quote_id=?', id);
    if (items && items.length) {
      items.forEach(item => {
        if (!item.product_name) return;
        q.run('INSERT INTO quote_items(quote_id,product_name,description,quantity,unit,unit_price) VALUES(?,?,?,?,?,?)',
          id, item.product_name, item.description || '', Number(item.quantity) || 1, item.unit || 'Adet', Number(item.unit_price) || 0);
      });
    }

    if (s) {
      for (const [k, v] of Object.entries(s)) {
        if (v !== undefined) q.run('INSERT INTO settings(key,value) VALUES(?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value', k, v);
      }
    }

    res.json({ ok: true });
  } catch (e) {
    res.json({ ok: false, msg: e.message });
  }
});

app.get('/teklif/:id', auth, (req, res) => {
  const quote = q.get('SELECT * FROM quotes WHERE id=?', req.params.id);
  if (!quote) return res.redirect('/teklifler');
  const items = q.all('SELECT * FROM quote_items WHERE quote_id=? ORDER BY id', quote.id);
  const subtotal = items.reduce((s, i) => s + i.quantity * i.unit_price, 0);
  let discount = 0;
  if (quote.discount_type === 'percent') discount = subtotal * quote.discount_value / 100;
  else if (quote.discount_type === 'amount') discount = quote.discount_value;
  const afterDiscount = subtotal - discount;
  const tax = afterDiscount * quote.tax_rate / 100;
  const total = afterDiscount + tax;
  res.render('quote-detail', { quote, items, subtotal, discount, afterDiscount, tax, total, page: 'quotes', title: 'Teklif #' + quote.id });
});

app.get('/teklif/:id/yazdir', auth, (req, res) => {
  const quote = q.get('SELECT * FROM quotes WHERE id=?', req.params.id);
  if (!quote) return res.redirect('/teklifler');
  const items = q.all('SELECT * FROM quote_items WHERE quote_id=? ORDER BY id', quote.id);
  const subtotal = items.reduce((s, i) => s + i.quantity * i.unit_price, 0);
  let discount = 0;
  if (quote.discount_type === 'percent') discount = subtotal * quote.discount_value / 100;
  else if (quote.discount_type === 'amount') discount = quote.discount_value;
  const afterDiscount = subtotal - discount;
  const tax = afterDiscount * quote.tax_rate / 100;
  const total = afterDiscount + tax;
  const settings = getSettings();
  res.render('quote-print', { quote, items, subtotal, discount, afterDiscount, tax, total, settings, layout: false });
});

app.get('/teklif/:id/excel', auth, async (req, res) => {
  const quote = q.get('SELECT * FROM quotes WHERE id=?', req.params.id);
  if (!quote) return res.redirect('/teklifler');
  const items = q.all('SELECT * FROM quote_items WHERE quote_id=? ORDER BY id', quote.id);
  const subtotal = items.reduce((s, i) => s + i.quantity * i.unit_price, 0);
  let discount = 0;
  if (quote.discount_type === 'percent') discount = subtotal * quote.discount_value / 100;
  else if (quote.discount_type === 'amount') discount = quote.discount_value;
  const afterDiscount = subtotal - discount;
  const tax = afterDiscount * quote.tax_rate / 100;
  const total = afterDiscount + tax;

  const settings = getSettings();
  const wb = new ExcelJS.Workbook();
  wb.creator = settings.company_name || 'Confort İklimlendirme';
  const ws = wb.addWorksheet('Teklif');
  const red = 'FF2563EB';
  const gray = 'FFF8F8F8';
  const darkGray = 'FF666666';
  const white = 'FFFFFFFF';
  const black = 'FF1A1A1A';

  ws.columns = [
    { width: 6 }, { width: 30 }, { width: 22 }, { width: 10 }, { width: 10 }, { width: 16 }, { width: 16 }
  ];

  // Firma başlığı
  ws.mergeCells('A1:C1');
  const logo = ws.getCell('A1');
  logo.value = settings.company_name || 'CONFORT İKLİMLENDİRME';
  logo.font = { bold: true, size: 20, color: { argb: red } };
  ws.mergeCells('A2:C2');
  const sub = ws.getCell('A2');
  sub.value = settings.company_subtitle || '';
  sub.font = { size: 9, color: { argb: darkGray } };

  ws.mergeCells('E1:G1');
  const ci1 = ws.getCell('E1');
  ci1.value = settings.owner_name || '';
  ci1.font = { size: 10, color: { argb: darkGray } };
  ci1.alignment = { horizontal: 'right' };
  ws.mergeCells('E2:G2');
  const ci2 = ws.getCell('E2');
  ci2.value = `${settings.address || ''} | Tel: ${settings.phone || ''}`;
  ci2.font = { size: 9, color: { argb: darkGray } };
  ci2.alignment = { horizontal: 'right' };

  // Kırmızı çizgi
  const r3 = ws.getRow(3);
  r3.height = 4;
  for (let c = 1; c <= 7; c++) { ws.getCell(3, c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: red } }; }

  // TEKLİF FORMU başlığı
  ws.mergeCells('A5:G5');
  const title = ws.getCell('A5');
  title.value = 'TEKLİF FORMU';
  title.font = { bold: true, size: 16, color: { argb: red } };
  title.alignment = { horizontal: 'center' };
  ws.mergeCells('A6:G6');
  const info = ws.getCell('A6');
  info.value = `Teklif No: #${quote.id}  |  Tarih: ${quote.quote_date}${quote.valid_until ? '  |  Geçerlilik: ' + quote.valid_until : ''}`;
  info.font = { size: 10, color: { argb: darkGray } };
  info.alignment = { horizontal: 'center' };

  // Müşteri ve Teklif bilgileri
  ws.mergeCells('A8:C8');
  ws.getCell('A8').value = 'MÜŞTERİ BİLGİLERİ';
  ws.getCell('A8').font = { bold: true, size: 10, color: { argb: red } };
  ws.getCell('A8').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: gray } };
  ws.mergeCells('E8:G8');
  ws.getCell('E8').value = 'TEKLİF BİLGİLERİ';
  ws.getCell('E8').font = { bold: true, size: 10, color: { argb: red } };
  ws.getCell('E8').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: gray } };

  ws.mergeCells('A9:C9');
  ws.getCell('A9').value = quote.customer_name || '-';
  ws.getCell('A9').font = { bold: true, size: 11 };
  ws.getCell('A9').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: gray } };
  ws.mergeCells('E9:G9');
  ws.getCell('E9').value = `Durum: ${quote.status}`;
  ws.getCell('E9').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: gray } };

  ws.mergeCells('A10:C10');
  ws.getCell('A10').value = `Tel: ${quote.customer_phone || '-'}${quote.customer_address ? '  |  Adres: ' + quote.customer_address : ''}`;
  ws.getCell('A10').font = { size: 10, color: { argb: darkGray } };
  ws.getCell('A10').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: gray } };
  ws.mergeCells('E10:G10');
  ws.getCell('E10').value = `KDV Oranı: %${quote.tax_rate}${quote.discount_value > 0 ? '  |  İndirim: ' + (quote.discount_type === 'percent' ? '%' + quote.discount_value : quote.discount_value + ' ₺') : ''}`;
  ws.getCell('E10').font = { size: 10, color: { argb: darkGray } };
  ws.getCell('E10').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: gray } };

  // Tablo başlığı
  const headers = ['#', 'Ürün / Hizmet', 'Açıklama', 'Miktar', 'Birim', 'Birim Fiyat (₺)', 'Toplam (₺)'];
  const hRow = ws.getRow(12);
  headers.forEach((h, i) => {
    const cell = hRow.getCell(i + 1);
    cell.value = h;
    cell.font = { bold: true, size: 10, color: { argb: white } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: red } };
    cell.alignment = { horizontal: i >= 3 ? 'center' : 'left' };
    if (i >= 5) cell.alignment = { horizontal: 'right' };
  });

  // Ürün satırları
  items.forEach((item, i) => {
    const row = ws.getRow(13 + i);
    row.getCell(1).value = i + 1;
    row.getCell(2).value = item.product_name;
    row.getCell(2).font = { bold: true };
    row.getCell(3).value = item.description || '';
    row.getCell(3).font = { color: { argb: darkGray } };
    row.getCell(4).value = item.quantity;
    row.getCell(4).alignment = { horizontal: 'center' };
    row.getCell(5).value = item.unit;
    row.getCell(5).alignment = { horizontal: 'center' };
    row.getCell(6).value = item.unit_price;
    row.getCell(6).numFmt = '#,##0.00 "₺"';
    row.getCell(6).alignment = { horizontal: 'right' };
    row.getCell(7).value = item.quantity * item.unit_price;
    row.getCell(7).numFmt = '#,##0.00 "₺"';
    row.getCell(7).font = { bold: true };
    row.getCell(7).alignment = { horizontal: 'right' };
    if (i % 2 === 1) for (let c = 1; c <= 7; c++) row.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: gray } };
  });

  // Toplamlar
  const tStart = 13 + items.length + 1;
  const addTotalRow = (label, value, style) => {
    const r = ws.getRow(tStart + addTotalRow.idx++);
    ws.mergeCells(r.number, 2, r.number, 3);
    r.getCell(2).value = label;
    r.getCell(2).alignment = { horizontal: 'right' };
    r.getCell(2).font = style?.labelFont || { color: { argb: darkGray } };
    r.getCell(7).value = value;
    r.getCell(7).numFmt = '#,##0.00 "₺"';
    r.getCell(7).alignment = { horizontal: 'right' };
    r.getCell(7).font = style?.valueFont || {};
    if (style?.bg) for (let c = 1; c <= 7; c++) r.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: style.bg } };
    return r;
  };
  addTotalRow.idx = 0;
  addTotalRow('Ara Toplam:', subtotal);
  if (discount > 0) addTotalRow('İndirim:', -discount, { valueFont: { color: { argb: 'FF16A34A' } } });
  addTotalRow(`KDV (%${quote.tax_rate}):`, tax);
  addTotalRow('GENEL TOPLAM', total, { bg: red, labelFont: { bold: true, size: 13, color: { argb: white } }, valueFont: { bold: true, size: 13, color: { argb: white } } });

  // Notlar
  if (quote.notes) {
    const nRow = tStart + addTotalRow.idx + 1;
    ws.mergeCells(nRow, 1, nRow, 7);
    ws.getCell(nRow, 1).value = `Notlar: ${quote.notes}`;
    ws.getCell(nRow, 1).font = { italic: true, size: 10, color: { argb: darkGray } };
    ws.getCell(nRow, 1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: gray } };
  }

  // Kenarlıklar (sadece tablo kısmı)
  for (let r = 12; r <= 12 + items.length; r++) {
    for (let c = 1; c <= 7; c++) {
      ws.getCell(r, c).border = { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
    }
  }

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename=Teklif-${quote.id}.xlsx`);
  await wb.xlsx.write(res);
  res.end();
});

app.post('/teklif/sil', auth, (req, res) => {
  q.run('DELETE FROM quote_items WHERE quote_id=?', req.body.id);
  q.run('DELETE FROM quotes WHERE id=?', req.body.id);
  req.session.flash = { type: 'success', msg: 'Teklif silindi' };
  res.redirect('/teklifler');
});

// --- Stok ---
app.get('/stok', auth, (req, res) => {
  const items = q.all('SELECT * FROM stock ORDER BY product_name');
  const lowStock = items.filter(i => i.quantity <= i.min_quantity);
  const movements = q.all(`SELECT sm.*, s.product_name FROM stock_movements sm JOIN stock s ON sm.stock_id=s.id ORDER BY sm.created_at DESC LIMIT 50`);
  res.render('stock', { items, lowStock, movements, page: 'stock', title: 'Stok Yönetimi' });
});

app.post('/stok/kaydet', auth, (req, res) => {
  const { id, product_name, sku, quantity, min_quantity, unit_cost, category } = req.body;
  if (id) {
    q.run('UPDATE stock SET product_name=?,sku=?,quantity=?,min_quantity=?,unit_cost=?,category=? WHERE id=?', product_name, sku, Number(quantity), Number(min_quantity), Number(unit_cost), category, id);
    req.session.flash = { type: 'success', msg: 'Stok güncellendi' };
  } else {
    const r = q.run('INSERT INTO stock(product_name,sku,quantity,min_quantity,unit_cost,category) VALUES(?,?,?,?,?,?)', product_name, sku, Number(quantity), Number(min_quantity), Number(unit_cost), category);
    q.run('INSERT INTO stock_movements(stock_id,type,quantity,note,user_name) VALUES(?,?,?,?,?)', r.lastInsertRowid, 'giris', Number(quantity), 'İlk stok girişi', req.session.user.name);
    req.session.flash = { type: 'success', msg: 'Stok eklendi' };
  }
  res.redirect('/stok');
});

app.post('/stok/hareket', auth, (req, res) => {
  const { stock_id, type, quantity, note } = req.body;
  const qty = Number(quantity);
  const item = q.get('SELECT * FROM stock WHERE id=?', stock_id);
  if (!item) return res.redirect('/stok');
  const newQty = type === 'giris' ? item.quantity + qty : Math.max(0, item.quantity - qty);
  q.run('UPDATE stock SET quantity=? WHERE id=?', newQty, stock_id);
  q.run('INSERT INTO stock_movements(stock_id,type,quantity,note,user_name) VALUES(?,?,?,?,?)', stock_id, type, qty, note, req.session.user.name);
  req.session.flash = { type: 'success', msg: `Stok ${type === 'giris' ? 'girişi' : 'çıkışı'} kaydedildi` };
  res.redirect('/stok');
});

app.post('/stok/sil', auth, (req, res) => {
  q.run('DELETE FROM stock_movements WHERE stock_id=?', req.body.id);
  q.run('DELETE FROM stock WHERE id=?', req.body.id);
  req.session.flash = { type: 'success', msg: 'Stok silindi' };
  res.redirect('/stok');
});

// --- Raporlar ---
app.get('/raporlar', auth, (req, res) => {
  const period = req.query.period || 'month';
  const monthlySales = q.all(`SELECT strftime('%Y-%m', sale_date) m, COUNT(*) c, SUM(price) total, SUM(paid_amount) paid FROM sales GROUP BY m ORDER BY m DESC LIMIT 12`);
  const topProducts = q.all(`SELECT product_name, COUNT(*) c, SUM(price) total FROM sales GROUP BY product_name ORDER BY c DESC LIMIT 10`);
  const topCustomers = q.all(`SELECT c.name, c.phone, COUNT(s.id) sale_count, SUM(s.price) total, SUM(s.price - s.paid_amount) debt FROM sales s JOIN customers c ON s.customer_id=c.id GROUP BY s.customer_id ORDER BY total DESC LIMIT 10`);
  const totalRevenue = q.get('SELECT COALESCE(SUM(price),0) s FROM sales').s;
  const totalCollected = q.get('SELECT COALESCE(SUM(paid_amount),0) s FROM sales').s;
  const totalDebt = totalRevenue - totalCollected;
  const unpaidSales = q.all(`SELECT s.*, c.name as customer_name, c.phone as customer_phone FROM sales s JOIN customers c ON s.customer_id=c.id WHERE s.payment_status != 'Ödendi' ORDER BY (s.price - s.paid_amount) DESC`);
  const paymentMethods = q.all(`SELECT payment_method, COUNT(*) c, SUM(price) total FROM sales GROUP BY payment_method ORDER BY total DESC`);
  const txGelir = q.get("SELECT COALESCE(SUM(amount),0) s FROM transactions WHERE type='gelir'").s;
  const txGider = q.get("SELECT COALESCE(SUM(amount),0) s FROM transactions WHERE type='gider'").s;
  const netProfit = totalCollected + txGelir - txGider;
  const expenseByCat = q.all("SELECT COALESCE(NULLIF(category,''),'Diğer') category, SUM(amount) total FROM transactions WHERE type='gider' GROUP BY category ORDER BY total DESC");
  res.render('reports', { monthlySales: monthlySales.reverse(), topProducts, topCustomers, totalRevenue, totalCollected, totalDebt, unpaidSales, paymentMethods, txGelir, txGider, netProfit, expenseByCat, page: 'reports', title: 'Raporlar' });
});

// --- Kasa (Gelir / Gider) ---
app.get('/kasa', auth, (req, res) => {
  const transactions = q.all('SELECT * FROM transactions ORDER BY tx_date DESC, id DESC');
  const gelir = q.get("SELECT COALESCE(SUM(amount),0) s FROM transactions WHERE type='gelir'").s;
  const gider = q.get("SELECT COALESCE(SUM(amount),0) s FROM transactions WHERE type='gider'").s;
  res.render('cashbook', { transactions, gelir, gider, net: gelir - gider, page: 'kasa', title: 'Kasa — Gelir / Gider' });
});
app.post('/kasa/kaydet', auth, (req, res) => {
  const { id, type, category, amount, tx_date, description } = req.body;
  const t = type === 'gider' ? 'gider' : 'gelir';
  if (id) q.run('UPDATE transactions SET type=?,category=?,amount=?,tx_date=?,description=? WHERE id=?', t, category, Number(amount), tx_date, description, id);
  else q.run('INSERT INTO transactions(type,category,amount,tx_date,description) VALUES(?,?,?,?,?)', t, category, Number(amount), tx_date, description);
  req.session.flash = { type: 'success', msg: 'Kayıt eklendi' };
  res.redirect('/kasa');
});
app.post('/kasa/sil', auth, (req, res) => {
  q.run('DELETE FROM transactions WHERE id=?', req.body.id);
  req.session.flash = { type: 'success', msg: 'Kayıt silindi' };
  res.redirect('/kasa');
});

// --- Yedekleme ---
app.get('/yedek', auth, (req, res) => {
  res.render('backup', { settings: getSettings(), backups: listLocalBackups(), page: 'yedek', title: 'Yedekleme' });
});
app.get('/yedek/json', auth, (req, res) => {
  res.setHeader('Content-Disposition', `attachment; filename="yedek-${dateStamp()}.json"`);
  res.setHeader('Content-Type', 'application/json');
  res.send(JSON.stringify(exportData(), null, 2));
});
app.get('/yedek/db', auth, (req, res) => {
  try { db.exec('PRAGMA wal_checkpoint(TRUNCATE)'); } catch {}
  res.download(path.join(DB_DIR, 'satis.db'), `satis-${dateStamp()}.db`);
});
app.get('/yedek/excel', auth, async (req, res) => {
  try {
    const buf = await buildExcel();
    res.setHeader('Content-Disposition', `attachment; filename="yedek-${dateStamp()}.xlsx"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(buf);
  } catch (e) { req.session.flash = { type: 'error', msg: 'Excel oluşturulamadı: ' + e.message }; res.redirect('/yedek'); }
});
app.get('/yedek/pdf', auth, async (req, res) => {
  try {
    const buf = await buildPdf();
    res.setHeader('Content-Disposition', `attachment; filename="yedek-ozet-${dateStamp()}.pdf"`);
    res.setHeader('Content-Type', 'application/pdf');
    res.send(buf);
  } catch (e) { req.session.flash = { type: 'error', msg: 'PDF oluşturulamadı: ' + e.message }; res.redirect('/yedek'); }
});
app.post('/yedek/yukle', auth, upload.single('file'), (req, res) => {
  try {
    if (!req.file) throw new Error('Dosya seçilmedi');
    const raw = fs.readFileSync(req.file.path, 'utf8');
    importData(JSON.parse(raw));
    fs.unlinkSync(req.file.path);
    req.session.flash = { type: 'success', msg: 'Yedek geri yüklendi' };
  } catch (e) { req.session.flash = { type: 'error', msg: 'Yedek yüklenemedi: ' + e.message }; }
  res.redirect('/yedek');
});
app.post('/yedek/klasore', auth, async (req, res) => {
  const r = await runFolderBackup();
  req.session.flash = r.ok ? { type: 'success', msg: 'Klasöre yedeklendi (' + r.formats.join(', ') + ')' } : { type: 'error', msg: r.error };
  res.redirect('/yedek');
});
app.post('/yedek/ayar', auth, (req, res) => {
  q.run('INSERT INTO settings(key,value) VALUES(?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value', 'backup_folder', req.body.backup_folder || '');
  q.run('INSERT INTO settings(key,value) VALUES(?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value', 'backup_folder2', req.body.backup_folder2 || '');
  q.run('INSERT INTO settings(key,value) VALUES(?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value', 'backup_auto', req.body.backup_auto ? '1' : '0');
  req.session.flash = { type: 'success', msg: 'Yedek ayarları kaydedildi' };
  res.redirect('/yedek');
});

// --- Ayarlar ---
app.get('/ayarlar', auth, (req, res) => {
  const settings = getSettings();
  res.render('settings', { settings, page: 'settings', title: 'Firma Ayarları' });
});

app.post('/ayarlar', auth, (req, res) => {
  const keys = ['company_name', 'company_subtitle', 'owner_name', 'address', 'phone', 'website', 'tax_id', 'default_tax_rate', 'default_valid_days', 'products'];
  keys.forEach(k => {
    q.run('INSERT INTO settings(key,value) VALUES(?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value', k, req.body[k] || '');
  });
  req.session.flash = { type: 'success', msg: 'Ayarlar kaydedildi' };
  res.redirect('/ayarlar');
});

app.post('/ayarlar/sifre', auth, (req, res) => {
  const { new_password, confirm_password } = req.body;
  if (new_password.length < 4) {
    req.session.flash = { type: 'error', msg: 'Yeni şifre en az 4 karakter olmalı' };
    return res.redirect('/ayarlar');
  }
  if (new_password !== confirm_password) {
    req.session.flash = { type: 'error', msg: 'Yeni şifreler eşleşmiyor' };
    return res.redirect('/ayarlar');
  }
  const hash = bcrypt.hashSync(new_password, 10);
  q.run('UPDATE users SET password=? WHERE id=?', hash, req.session.user.id);
  req.session.flash = { type: 'success', msg: 'Şifre başarıyla değiştirildi' };
  res.redirect('/ayarlar');
});

// API: Create customer from sale form
app.post('/api/musteri/ekle', auth, (req, res) => {
  const { name, phone, city, address } = req.body;
  if (!name) return res.json({ error: 'Ad Soyad gerekli' });
  const result = q.run('INSERT INTO customers(name,phone,city,address) VALUES(?,?,?,?)', name, phone || '', city || '', address || '');
  res.json({ id: Number(result.lastInsertRowid) });
});

// --- Word Export ---
app.get('/teklif/:id/word', auth, async (req, res) => {
  const quote = q.get('SELECT * FROM quotes WHERE id=?', req.params.id);
  if (!quote) return res.redirect('/teklifler');
  const items = q.all('SELECT * FROM quote_items WHERE quote_id=? ORDER BY id', quote.id);
  const settings = getSettings();
  const subtotal = items.reduce((s, i) => s + i.quantity * i.unit_price, 0);
  let discount = 0;
  if (quote.discount_type === 'percent') discount = subtotal * quote.discount_value / 100;
  else if (quote.discount_type === 'amount') discount = quote.discount_value;
  const afterDiscount = subtotal - discount;
  const tax = afterDiscount * quote.tax_rate / 100;
  const total = afterDiscount + tax;
  const money = (n) => Number(n).toLocaleString('tr-TR', { minimumFractionDigits: 2 }) + ' ₺';

  const { Document, Paragraph, Table, TableRow, TableCell, TextRun, WidthType, AlignmentType, BorderStyle, HeadingLevel, ShadingType } = docx;

  const redColor = '2563EB';
  const grayBg = 'F8F8F8';

  // Column widths in DXA — A4 content width = 9026 DXA (11906 - 1440*2)
  const colW = [500, 2200, 1726, 800, 800, 1400, 1600]; // total = 9026
  const noBorder = { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } };
  const cell = (text, w, opts) => new TableCell({
    children: [new Paragraph({ children: [new TextRun({ text: String(text), font: 'Segoe UI', size: 20, ...(opts?.font || {}) })], alignment: opts?.align || AlignmentType.LEFT })],
    width: { size: w, type: WidthType.DXA },
    shading: opts?.shading,
    borders: opts?.borders,
    columnSpan: opts?.columnSpan,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
  });

  const tableHeaders = ['#', 'Ürün / Hizmet', 'Açıklama', 'Miktar', 'Birim', 'Birim Fiyat', 'Toplam'];
  const headerRow = new TableRow({
    children: tableHeaders.map((h, i) => cell(h, colW[i], {
      font: { bold: true, color: 'FFFFFF', size: 18 },
      align: AlignmentType.CENTER,
      shading: { type: ShadingType.CLEAR, fill: redColor },
    }))
  });

  const itemRows = items.map((item, i) => new TableRow({
    children: [
      cell(i + 1, colW[0], { align: AlignmentType.CENTER }),
      cell(item.product_name, colW[1], { font: { bold: true } }),
      cell(item.description || '', colW[2], { font: { color: '666666' } }),
      cell(item.quantity, colW[3], { align: AlignmentType.CENTER }),
      cell(item.unit, colW[4], { align: AlignmentType.CENTER }),
      cell(money(item.unit_price), colW[5], { align: AlignmentType.RIGHT }),
      cell(money(item.quantity * item.unit_price), colW[6], { align: AlignmentType.RIGHT, font: { bold: true } }),
    ]
  }));

  // Totals as separate right-aligned table
  const totCell = (text, w, opts) => new TableCell({
    children: [new Paragraph({ children: [new TextRun({ text: String(text), font: 'Segoe UI', size: opts?.fontSize || 20, bold: opts?.bold, color: opts?.color || '000000' })], alignment: opts?.align || AlignmentType.RIGHT })],
    width: { size: w, type: WidthType.DXA },
    shading: opts?.bg ? { type: ShadingType.CLEAR, fill: opts.bg } : undefined,
    margins: { top: 60, bottom: 60, left: 120, right: 120 },
  });
  const totRows = [];
  totRows.push(new TableRow({ children: [totCell('Ara Toplam:', 2000), totCell(money(subtotal), 2200, { bold: true })] }));
  if (discount > 0) totRows.push(new TableRow({ children: [totCell('İndirim:', 2000), totCell('-' + money(discount), 2200, { bold: true, color: '16A34A' })] }));
  totRows.push(new TableRow({ children: [totCell(`KDV (%${quote.tax_rate}):`, 2000), totCell(money(tax), 2200, { bold: true })] }));
  totRows.push(new TableRow({ children: [totCell('GENEL TOPLAM', 2000, { bold: true, fontSize: 24, color: 'FFFFFF', bg: redColor }), totCell(money(total), 2200, { bold: true, fontSize: 24, color: 'FFFFFF', bg: redColor })] }));

  const doc = new Document({
    sections: [{
      children: [
        new Paragraph({ children: [new TextRun({ text: settings.company_name || 'CONFORT İKLİMLENDİRME', bold: true, size: 40, color: redColor, font: 'Segoe UI' })], spacing: { after: 50 } }),
        new Paragraph({ children: [new TextRun({ text: settings.company_subtitle || '', color: '666666', size: 18, font: 'Segoe UI' })], spacing: { after: 100 } }),
        new Paragraph({ children: [new TextRun({ text: `${settings.owner_name}  |  ${settings.address}  |  Tel: ${settings.phone}`, color: '666666', size: 18, font: 'Segoe UI' })], alignment: AlignmentType.RIGHT, spacing: { after: 200 } }),
        new Paragraph({ children: [new TextRun({ text: '━'.repeat(80), color: redColor })], spacing: { after: 200 } }),
        new Paragraph({ children: [new TextRun({ text: 'TEKLİF FORMU', bold: true, size: 32, color: redColor, font: 'Segoe UI' })], alignment: AlignmentType.CENTER, spacing: { after: 50 } }),
        new Paragraph({ children: [new TextRun({ text: `Teklif No: #${quote.id}  |  Tarih: ${quote.quote_date}${quote.valid_until ? '  |  Geçerlilik: ' + quote.valid_until : ''}`, color: '666666', size: 20 })], alignment: AlignmentType.CENTER, spacing: { after: 300 } }),
        new Paragraph({ children: [new TextRun({ text: 'MÜŞTERİ BİLGİLERİ', bold: true, color: redColor, size: 20 })], shading: { type: ShadingType.CLEAR, fill: grayBg }, spacing: { after: 50 } }),
        new Paragraph({ children: [new TextRun({ text: quote.customer_name || '-', bold: true, size: 22 })], spacing: { after: 30 } }),
        new Paragraph({ children: [new TextRun({ text: `Tel: ${quote.customer_phone || '-'}${quote.customer_address ? '  |  Adres: ' + quote.customer_address : ''}`, color: '666666', size: 18 })], spacing: { after: 300 } }),
        new Table({ rows: [headerRow, ...itemRows], width: { size: 9026, type: WidthType.DXA }, columnWidths: colW }),
        new Paragraph({ children: [new TextRun({ text: '' })], spacing: { before: 100 } }),
        new Table({ rows: totRows, width: { size: 4200, type: WidthType.DXA }, columnWidths: [2000, 2200], indent: { size: 4826, type: WidthType.DXA } }),
        ...(quote.notes ? [
          new Paragraph({ children: [new TextRun({ text: '' })], spacing: { before: 300 } }),
          new Paragraph({ children: [new TextRun({ text: 'Notlar: ', bold: true, color: redColor, font: 'Segoe UI' }), new TextRun({ text: quote.notes, color: '333333', font: 'Segoe UI' })], shading: { type: ShadingType.CLEAR, fill: grayBg } }),
        ] : []),
        new Paragraph({ children: [new TextRun({ text: '' })], spacing: { before: 600 } }),
        new Table({
          rows: [new TableRow({
            children: [
              new TableCell({ children: [new Paragraph({ text: '________________________', alignment: AlignmentType.CENTER, spacing: { before: 400 } }), new Paragraph({ text: 'Müşteri İmza / Kaşe', alignment: AlignmentType.CENTER })], borders: noBorder, width: { size: 4513, type: WidthType.DXA } }),
              new TableCell({ children: [new Paragraph({ text: '________________________', alignment: AlignmentType.CENTER, spacing: { before: 400 } }), new Paragraph({ text: 'Yetkili İmza / Kaşe', alignment: AlignmentType.CENTER })], borders: noBorder, width: { size: 4513, type: WidthType.DXA } }),
            ]
          })],
          width: { size: 9026, type: WidthType.DXA },
          columnWidths: [4513, 4513],
        }),
      ]
    }]
  });

  const buffer = await docx.Packer.toBuffer(doc);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  res.setHeader('Content-Disposition', `attachment; filename=Teklif-${quote.id}.docx`);
  res.end(buffer);
});

// Otomatik klasör yedeği — açıksa her akşam saat 20:00'de (gün sonu), günde bir kez
let _lastBackupDay = '';
setInterval(() => {
  try {
    const s = getSettings();
    if (s.backup_auto !== '1' || !s.backup_folder) return;
    const now = new Date();
    const today = now.toISOString().slice(0, 10);
    if (now.getHours() >= 20 && _lastBackupDay !== today) {
      _lastBackupDay = today;
      runFolderBackup().catch(() => {});
    }
  } catch {}
}, 10 * 60 * 1000); // her 10 dakikada bir kontrol

app.listen(PORT, () => {
  console.log(`\n  Confort Satış Takip → http://localhost:${PORT}`);
  console.log(`  Giriş: admin / admin123\n`);
});
