import "dotenv/config";
import fetch from 'node-fetch';
import { createServer } from 'http';
import { parse } from 'querystring';

const PORT = process.env.PORT || 3000;
const HOST = '127.0.0.1';
const APP_TITLE = 'Monchai Insurance';
const SHEET_WRITE_URL = process.env.SHEET_WEBHOOK_URL;
const SHEET_DATA_URL = process.env.SHEET_DATA_URL;
const SHEET_UPDATE_URL = process.env.SHEET_UPDATE_URL || SHEET_WRITE_URL;
const DEFAULT_EXPIRY_WINDOW_DAYS = 30;

function escapeHtml(value = '') {
  const text = value == null ? '' : String(value);
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function formatDate(dateIso) {
  if (!dateIso) return '-';
  const date = new Date(dateIso);
  if (Number.isNaN(date.getTime())) return '-';
  return date.toLocaleDateString('th-TH', { year: 'numeric', month: 'short', day: 'numeric' });
}

function toInputDate(value) {
  if (value == null || value === '') return '';
  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(trimmed)) return trimmed;
    const parsed = Date.parse(trimmed);
    if (Number.isNaN(parsed)) return '';
    return toInputDate(new Date(parsed));
  }
  const date = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(date.getTime())) return '';
  const local = new Date(date.getTime() - date.getTimezoneOffset() * 60000);
  return local.toISOString().slice(0, 10);
}

function renderLayout({ pageTitle, active = '', content = '', alertMarkup = '' }) {
  const navItems = [
    { href: '/', label: 'หน้าหลัก', key: 'home' },
    { href: '/customers/new', label: 'เพิ่มลูกค้าใหม่', key: 'new' },
    { href: '/customers/search', label: 'ค้นหาลูกค้า', key: 'search' },
    { href: '/customers/expiring', label: 'แจ้งเตือนลูกค้าที่จะหมดอายุ', key: 'expiring' }
  ];

  const navMarkup = navItems
    .map(item => {
      const isActive = item.key === active;
      return `<a class="nav-link${isActive ? ' active' : ''}" href="${item.href}">${item.label}</a>`;
    })
    .join('');

  return `<!DOCTYPE html>
<html lang="th">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>${escapeHtml(pageTitle)} - ${APP_TITLE}</title>
  <style>
    :root { color-scheme: light; }
    body { margin: 0; min-height: 100vh; display: flex; flex-direction: column; background: linear-gradient(155deg, #0f172a 0%, #1d3359 45%, #245796 100%); font-family: 'Segoe UI', Tahoma, Arial, sans-serif; color: #1f2933; }
    header { background: rgba(15, 23, 42, 0.85); color: #f8fafc; padding: 1.25rem 2rem; display: flex; align-items: center; justify-content: space-between; gap: 1rem; }
    .brand { display: flex; align-items: center; gap: 0.8rem; }
    .brand-badge { width: 46px; height: 46px; border-radius: 14px; background: linear-gradient(135deg, #2b6cb0, #63b3ed); display: flex; align-items: center; justify-content: center; color: #ffffff; font-weight: 700; font-size: 1.15rem; letter-spacing: 0.06em; box-shadow: 0 10px 24px rgba(43, 108, 176, 0.35); }
    header h1 { margin: 0; font-size: 1.4rem; font-weight: 600; }
    nav { display: flex; flex-wrap: wrap; gap: 0.6rem; }
    .nav-link { padding: 0.45rem 0.9rem; border-radius: 999px; border: 1px solid rgba(148, 163, 184, 0.5); color: #e2e8f0; text-decoration: none; font-size: 0.95rem; transition: background 0.2s ease, color 0.2s ease, border 0.2s ease; }
    .nav-link:hover, .nav-link:focus { background: rgba(148, 163, 184, 0.2); border-color: rgba(226, 232, 240, 0.7); }
    .nav-link.active { background: #2b6cb0; border-color: #2b6cb0; color: #ffffff; box-shadow: 0 12px 24px rgba(43, 108, 176, 0.35); }
    main { flex: 1 1 auto; display: flex; justify-content: center; padding: 2.5rem 1.5rem 3rem; }
    .card { background: #ffffff; border-radius: 20px; width: min(1024px, 100%); padding: 2.5rem; box-shadow: 0 30px 60px rgba(9, 16, 31, 0.25); }
    .card h2 { margin-top: 0; font-size: 1.75rem; color: #102a43; }
    .card p.lead { margin: 0.4rem 0 1.8rem; font-size: 1rem; color: #486581; }
    .alert { margin-bottom: 1.5rem; padding: 0.9rem 1rem; border-radius: 12px; font-size: 0.95rem; border: 1px solid transparent; }
    .alert.success { background: #ebf8ff; border-color: #90cdf4; color: #1a365d; }
    .alert.error { background: #fff5f5; border-color: #feb2b2; color: #742a2a; }
    form.grid { display: grid; gap: 1rem; }
    @media (min-width: 540px) { form.grid { grid-template-columns: repeat(2, minmax(0, 1fr)); } form.grid .full-width { grid-column: 1 / -1; } }
    .field { display: flex; flex-direction: column; gap: 0.45rem; }
    label { font-weight: 600; font-size: 0.95rem; color: #243b53; }
    input[type="text"], input[type="tel"], input[type="date"], textarea, select { padding: 0.8rem 0.95rem; font-size: 1rem; border-radius: 10px; border: 1px solid #d9e2ec; background: #f8fafc; transition: border-color 0.2s ease, box-shadow 0.2s ease; resize: vertical; min-height: 48px; }
    textarea { min-height: 96px; }
    input:focus, textarea:focus, select:focus { border-color: #2c5282; box-shadow: 0 0 0 4px rgba(44, 82, 130, 0.15); outline: none; }
    input.invalid, textarea.invalid { border-color: #e53e3e; box-shadow: 0 0 0 4px rgba(229, 62, 62, 0.1); }
    .field-error { color: #c53030; font-size: 0.85rem; margin: -0.35rem 0 0; }
    button.primary { padding: 0.9rem 1rem; font-size: 1rem; font-weight: 600; border-radius: 10px; border: none; cursor: pointer; color: #ffffff; background: linear-gradient(135deg, #2b6cb0, #26519a); box-shadow: 0 16px 30px rgba(37, 99, 235, 0.25); transition: transform 0.15s ease, box-shadow 0.15s ease; }
    button.primary:hover { transform: translateY(-1px); box-shadow: 0 18px 36px rgba(37, 99, 235, 0.3); }
    button.primary:active { transform: translateY(0); box-shadow: 0 10px 26px rgba(37, 99, 235, 0.25); }
    .table { width: 100%; border-collapse: collapse; margin-top: 1rem; }
    .table thead { background: rgba(43, 108, 176, 0.08); }
    .table th, .table td { padding: 0.75rem; border-bottom: 1px solid #e1e8f0; text-align: left; font-size: 0.95rem; vertical-align: top; }
    .table tbody tr:hover { background: rgba(148, 163, 184, 0.12); }
    .table a.record-link { color: #2c5282; text-decoration: none; font-weight: 600; }
    .table a.record-link:hover { text-decoration: underline; }
    .muted { color: #627d98; font-size: 0.9rem; }
    footer { text-align: center; padding: 1rem 0 2rem; color: rgba(255, 255, 255, 0.7); font-size: 0.85rem; }
    /* Home menu (icon grid) */
    .menu-grid { display: grid; grid-template-columns: repeat(1, minmax(0, 1fr)); gap: 1rem; }
    @media (min-width: 520px) { .menu-grid { grid-template-columns: repeat(2, minmax(0, 1fr)); } }
    @media (min-width: 900px) { .menu-grid { grid-template-columns: repeat(3, minmax(0, 1fr)); } }
    .menu-card { display: flex; flex-direction: column; gap: 0.5rem; padding: 1rem; border-radius: 16px; text-decoration: none; background: #f8fafc; border: 1px solid #e1e8f0; color: #102a43; transition: transform 0.12s ease, box-shadow 0.12s ease, background 0.12s ease, border-color 0.12s ease; }
    .menu-card:hover { transform: translateY(-2px); background: #ffffff; box-shadow: 0 14px 28px rgba(15, 23, 42, 0.15); border-color: #d5e0ea; }
    .menu-card strong { font-size: 1.05rem; }
    .menu-card span { color: #486581; font-size: 0.95rem; }
    .menu-icon { width: 48px; height: 48px; display: inline-flex; align-items: center; justify-content: center; border-radius: 12px; background: linear-gradient(135deg, #2b6cb0, #63b3ed); color: #ffffff; font-size: 1.25rem; box-shadow: 0 10px 24px rgba(43, 108, 176, 0.35); }
  </style>
</head>
<body>
  <header>
    <div class="brand">
      <div class="brand-badge">MI</div>
      <div>
        <h1>${APP_TITLE}</h1>
        <div class="muted">วางแผนติดต่อลูกค้าเพื่อการต่ออายุกรมธรรม์</div>
      </div>
    </div>
    <nav>${navMarkup}</nav>
  </header>
  <main>
    <article class="card">
      ${alertMarkup}
      ${content}
    </article>
  </main>
  <footer>Monchai Insurance · ระบบช่วยเตือนต่ออายุลูกค้า</footer>
</body>
</html>`;
}

function renderHomePage() {
  const content = `
    <h2>ภาพรวมการทำงาน</h2>
    <p class="lead">เลือกเมนูที่ต้องการจัดการลูกค้า เริ่มจากบันทึกลูกค้าใหม่ ไปจนถึงค้นหาและดูรายการที่ใกล้หมดอายุ</p>
    <div class="menu-grid">
      <a class="menu-card" href="/customers/new">
        <div class="menu-icon">📝</div>
        <strong>เพิ่มลูกค้าใหม่</strong>
        <span>กรอกข้อมูลลูกค้าและตั้งบันทึกการติดต่อล่วงหน้า</span>
      </a>
      <a class="menu-card" href="/customers/search">
        <div class="menu-icon">🔍</div>
        <strong>ค้นหาลูกค้า</strong>
        <span>ดูข้อมูลลูกค้าทั้งหมดและค้นหาตามชื่อหรือทะเบียนรถ</span>
      </a>
      <a class="menu-card" href="/customers/expiring">
        <div class="menu-icon">⏰</div>
        <strong>แจ้งเตือนลูกค้าที่จะหมดอายุ</strong>
        <span>ดูรายชื่อลูกค้าที่มีกำหนดต่ออายุภายใน ${DEFAULT_EXPIRY_WINDOW_DAYS} วัน</span>
      </a>
    </div>
  `;
  return renderLayout({ pageTitle: 'หน้าหลัก', active: 'home', content });
}

function renderCustomerForm({
  heading,
  lead,
  action,
  submitLabel,
  formData = {},
  errors = {},
  message = '',
  status = '',
  includeHidden = ''
} = {}) {
  const safe = key => escapeHtml(formData[key] ?? '');
  const valueFor = key => escapeHtml(formData[`${key}Input`] ?? formData[key] ?? '');
  const alertMarkup = message ? `<div class="alert ${status}">${escapeHtml(message)}</div>` : '';
  const fieldError = field => errors[field] ? `<p class="field-error">${escapeHtml(errors[field])}</p>` : '';

  const content = `
    <h2>${escapeHtml(heading)}</h2>
    <p class="lead">${escapeHtml(lead)}</p>
    <form class="grid" method="POST" action="${escapeHtml(action)}" novalidate>
      ${includeHidden}
      <div class="field full-width">
        <label for="customerName">ชื่อลูกค้า</label>
        <input id="customerName" name="customerName" type="text" maxlength="100" placeholder="ระบุชื่อ-นามสกุลลูกค้า" value="${safe('customerName')}" required class="${errors.customerName ? 'invalid' : ''}" />
        ${fieldError('customerName')}
      </div>
      <div class="field">
        <label for="licensePlate">ทะเบียนรถ</label>
        <input id="licensePlate" name="licensePlate" type="text" maxlength="60" placeholder="เช่น 1กก-1234" value="${safe('licensePlate')}" required class="${errors.licensePlate ? 'invalid' : ''}" />
        ${fieldError('licensePlate')}
      </div>
      <div class="field">
        <label for="phone">เบอร์ติดต่อหลัก</label>
        <input id="phone" name="phone" type="tel" maxlength="40" placeholder="เช่น 081-234-5678" value="${safe('phone')}" required class="${errors.phone ? 'invalid' : ''}" />
        ${fieldError('phone')}
      </div>
      <div class="field">
        <label for="actIssuedDate">วันที่ทำ พ.ร.บ.</label>
        <input id="actIssuedDate" name="actIssuedDate" type="date" value="${valueFor('actIssuedDate')}" class="${errors.actIssuedDate ? 'invalid' : ''}" />
        ${fieldError('actIssuedDate')}
      </div>
      <div class="field">
        <label for="actExpiryDate">วันที่ครบกำหนด พ.ร.บ.</label>
        <input id="actExpiryDate" name="actExpiryDate" type="date" value="${valueFor('actExpiryDate')}" class="${errors.actExpiryDate ? 'invalid' : ''}" />
        ${fieldError('actExpiryDate')}
      </div>
      <div class="field">
        <label for="taxRenewalDate">วันที่ต่อภาษี</label>
        <input id="taxRenewalDate" name="taxRenewalDate" type="date" value="${valueFor('taxRenewalDate')}" class="${errors.taxRenewalDate ? 'invalid' : ''}" />
        ${fieldError('taxRenewalDate')}
      </div>
      <div class="field">
        <label for="taxExpiryDate">วันที่ครบกำหนดต่อภาษี</label>
        <input id="taxExpiryDate" name="taxExpiryDate" type="date" value="${valueFor('taxExpiryDate')}" class="${errors.taxExpiryDate ? 'invalid' : ''}" />
        ${fieldError('taxExpiryDate')}
      </div>
      <div class="field">
        <label for="voluntaryIssuedDate">วันที่ทำกรมธรรม์ภาคสมัครใจ</label>
        <input id="voluntaryIssuedDate" name="voluntaryIssuedDate" type="date" value="${valueFor('voluntaryIssuedDate')}" class="${errors.voluntaryIssuedDate ? 'invalid' : ''}" />
        ${fieldError('voluntaryIssuedDate')}
      </div>
      <div class="field">
        <label for="voluntaryExpiryDate">วันที่ครบกำหนดกรมธรรม์ภาคสมัครใจ</label>
        <input id="voluntaryExpiryDate" name="voluntaryExpiryDate" type="date" value="${valueFor('voluntaryExpiryDate')}" class="${errors.voluntaryExpiryDate ? 'invalid' : ''}" />
        ${fieldError('voluntaryExpiryDate')}
      </div>
      <div class="field full-width">
        <label for="notes">หมายเหตุ</label>
        <textarea id="notes" name="notes" maxlength="500" placeholder="สรุปสิ่งที่ต้องพูดคุย ประวัติการติดตาม หรือข้อเสนอพิเศษสำหรับลูกค้า">${safe('notes')}</textarea>
      </div>
      <div class="field full-width">
        <button class="primary" type="submit">${escapeHtml(submitLabel)}</button>
      </div>
    </form>
  `;

  const activeTab = action.includes('/customers/update') ? 'search' : 'new';
  return renderLayout({ pageTitle: heading, active: activeTab, content, alertMarkup });
}

function renderAddCustomerPage(options = {}) {
  return renderCustomerForm({
    heading: 'เพิ่มลูกค้าใหม่',
    lead: 'บันทึกข้อมูลลูกค้าเพื่อเตรียมการติดต่อแจ้งต่ออายุและงานเอกสารต่าง ๆ',
    action: '/customers',
    submitLabel: 'บันทึกข้อมูลลูกค้า',
    ...options
  });
}

function renderEditCustomerPage(options = {}) {
  const { formData = {} } = options;
  const hidden = `
    <input type="hidden" name="rowNumber" value="${escapeHtml(formData.rowNumber ?? '')}" />
    <input type="hidden" name="timestamp" value="${escapeHtml(formData.timestamp ?? '')}" />
  `;
  return renderCustomerForm({
    heading: 'แก้ไขข้อมูลลูกค้า',
    lead: 'อัปเดตข้อมูลเพื่อติดตามงานต่ออายุและแจ้งเตือนได้อย่างแม่นยำ',
    action: '/customers/update',
    submitLabel: 'บันทึกการแก้ไข',
    includeHidden: hidden,
    formData,
    ...options
  });
}

function renderSearchPage({ query = '', results = [], total = 0 }) {
  const safeQuery = escapeHtml(query);
  const infoMarkup = `<p class="muted">พบลูกค้าทั้งหมด ${total} ราย${query ? ` · กรองจากคำค้นหา "${safeQuery}"` : ''}</p>`;

  const describePair = (labelA, dateA, labelB, dateB) => {
    return `<div><strong>${labelA}:</strong> ${escapeHtml(formatDate(dateA))}</div>
            <div><strong>${labelB}:</strong> ${escapeHtml(formatDate(dateB))}</div>`;
  };

  const tableContent = results.length
    ? `<table class="table">
        <thead>
          <tr>
            <th>ชื่อลูกค้า</th>
            <th>ทะเบียนรถ</th>
            <th>พ.ร.บ.</th>
            <th>ต่อภาษี</th>
            <th>ภาคสมัครใจ</th>
            <th>เบอร์ติดต่อ</th>
            <th>หมายเหตุ</th>
          </tr>
        </thead>
        <tbody>
          ${results.map(record => `
            <tr>
              <td><a class="record-link" href="/customers/edit?row=${record.rowNumber ?? ''}">${escapeHtml(record.customerName || '')}</a></td>
              <td>${escapeHtml(record.licensePlate || '')}</td>
              <td>${describePair('ทำ', record.actIssuedDate, 'ครบกำหนด', record.actExpiryDate)}</td>
              <td>${describePair('ต่อ', record.taxRenewalDate, 'ครบกำหนด', record.taxExpiryDate)}</td>
              <td>${describePair('ทำ', record.voluntaryIssuedDate, 'ครบกำหนด', record.voluntaryExpiryDate)}</td>
              <td>${escapeHtml(record.phone || '')}</td>
              <td>${escapeHtml(record.notes || '')}</td>
            </tr>`).join('')}
        </tbody>
      </table>`
    : '<p class="muted">ยังไม่มีข้อมูลลูกค้าให้แสดง</p>';

  const content = `
    <h2>ค้นหาลูกค้า</h2>
    <p class="lead">ค้นหาตามชื่อลูกค้า หรือทะเบียนรถ ข้อมูลทั้งหมดจะแสดงด้านล่าง</p>
    <form method="GET" action="/customers/search" class="grid">
      <div class="field full-width">
        <label for="q">คำค้นหา</label>
        <input id="q" name="q" type="text" value="${safeQuery}" placeholder="กรอกชื่อลูกค้า หรือทะเบียนรถ" />
      </div>
      <div class="field">
        <button class="primary" type="submit">ค้นหา</button>
      </div>
    </form>
    ${infoMarkup}
    ${tableContent}
  `;
  return renderLayout({ pageTitle: 'ค้นหาลูกค้า', active: 'search', content });
}

function renderExpiringPage({ customers = [], days = DEFAULT_EXPIRY_WINDOW_DAYS }) {
  const describePair = (labelA, dateA, labelB, dateB) => {
    return `<div><strong>${labelA}:</strong> ${escapeHtml(formatDate(dateA))}</div>
            <div><strong>${labelB}:</strong> ${escapeHtml(formatDate(dateB))}</div>`;
  };

  const tableContent = customers.length
    ? `<table class="table">
        <thead>
          <tr>
            <th>ชื่อลูกค้า</th>
            <th>ทะเบียนรถ</th>
            <th>พ.ร.บ.</th>
            <th>เหลืออีก (วัน)</th>
            <th>เบอร์ติดต่อ</th>
            <th>หมายเหตุ</th>
          </tr>
        </thead>
        <tbody>
          ${customers.map(item => `
            <tr>
              <td><a class="record-link" href="/customers/edit?row=${item.customer.rowNumber ?? ''}">${escapeHtml(item.customer.customerName || '')}</a></td>
              <td>${escapeHtml(item.customer.licensePlate || '')}</td>
              <td>${describePair('ทำ', item.customer.actIssuedDate, 'ครบกำหนด', item.customer.actExpiryDate)}</td>
              <td>${item.daysRemaining}</td>
              <td>${escapeHtml(item.customer.phone || '')}</td>
              <td>${escapeHtml(item.customer.notes || '')}</td>
            </tr>`).join('')}
        </tbody>
      </table>`
    : `<p class="muted">ไม่มีลูกค้าที่จะหมดอายุภายใน ${days} วัน</p>`;

  const content = `
    <h2>แจ้งเตือนลูกค้าที่จะหมดอายุ</h2>
    <p class="lead">ดูรายชื่อลูกค้าที่มีกำหนดต่ออายุภายในช่วง ${days} วันข้างหน้า (อ้างอิงจากวันที่ครบกำหนด พ.ร.บ.)</p>
    <form method="GET" action="/customers/expiring" class="grid">
      <div class="field">
        <label for="days">ระบุช่วงวันล่วงหน้า</label>
        <input id="days" name="days" type="number" min="1" max="365" value="${days}" />
      </div>
      <div class="field">
        <button class="primary" type="submit">อัปเดตรายการ</button>
      </div>
    </form>
    ${tableContent}
  `;
  return renderLayout({ pageTitle: 'แจ้งเตือนลูกค้าที่จะหมดอายุ', active: 'expiring', content });
}

async function syncToGoogleSheet(record) {
  if (!SHEET_WRITE_URL) {
    console.warn('SHEET_WEBHOOK_URL is not set; skipping Google Sheet sync.');
    return { ok: false, reason: 'missing-write-url' };
  }
  try {
    const response = await fetch(SHEET_WRITE_URL, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(record) });
    if (!response.ok) return { ok: false, reason: 'response-error', status: response.status, text: await response.text() };
    return { ok: true };
  } catch (error) {
    console.error('Google Sheet sync error:', error);
    return { ok: false, reason: 'exception', error };
  }
}

async function updateCustomerInSheet(rowNumber, record) {
  if (!rowNumber) return { ok: false, reason: 'missing-row-number' };
  if (!SHEET_UPDATE_URL) return { ok: false, reason: 'missing-update-url' };
  const payload = { action: 'update', rowNumber, record };
  try {
    const response = await fetch(SHEET_UPDATE_URL, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) });
    if (!response.ok) return { ok: false, reason: 'response-error', status: response.status, text: await response.text() };
    return { ok: true };
  } catch (error) {
    console.error('Google Sheet update error:', error);
    return { ok: false, reason: 'exception', error };
  }
}

function normaliseRecord(raw = {}, index = 0) {
  const base = {
    timestamp: raw.timestamp || raw.Timestamp || raw['Timestamp'] || null,
    customerName: raw.customerName || raw.CustomerName || raw['ชื่อลูกค้า'] || '',
    licensePlate: raw.licensePlate || raw.LicensePlate || raw['ทะเบียนรถ'] || '',
    actIssuedDate: raw.actIssuedDate || raw.ActIssuedDate || raw['วันที่ทำ พ.ร.บ.'] || '',
    actExpiryDate: raw.actExpiryDate || raw.ActExpiryDate || raw['วันที่ครบกำหนด พ.ร.บ.'] || '',
    taxRenewalDate: raw.taxRenewalDate || raw.TaxRenewalDate || raw['วันที่ต่อภาษี'] || '',
    taxExpiryDate: raw.taxExpiryDate || raw.TaxExpiryDate || raw['วันที่ครบกำหนดต่อภาษี'] || '',
    voluntaryIssuedDate: raw.voluntaryIssuedDate || raw.VoluntaryIssuedDate || raw['วันที่ทำกรมธรรม์ภาคสมัครใจ'] || '',
    voluntaryExpiryDate: raw.voluntaryExpiryDate || raw.VoluntaryExpiryDate || raw['วันที่ครบกำหนดกรมธรรม์ภาคสมัครใจ'] || '',
    phone: raw.phone || raw.Phone || raw['เบอร์ติดต่อหลัก'] || '',
    notes: raw.notes || raw.Notes || raw['หมายเหตุ'] || raw['บันทึก'] || ''
  };
  const rowNumber = raw.rowNumber || raw.row || raw.__rowNumber || raw.__row || (index + 2);
  return {
    ...base,
    rowNumber,
    actIssuedDateInput: toInputDate(base.actIssuedDate),
    actExpiryDateInput: toInputDate(base.actExpiryDate),
    taxRenewalDateInput: toInputDate(base.taxRenewalDate),
    taxExpiryDateInput: toInputDate(base.taxExpiryDate),
    voluntaryIssuedDateInput: toInputDate(base.voluntaryIssuedDate),
    voluntaryExpiryDateInput: toInputDate(base.voluntaryExpiryDate)
  };
}

function getRecordSortTime(record) {
  const value = record.actExpiryDate || record.timestamp || record.actIssuedDate || '';
  const time = new Date(value).getTime();
  return Number.isNaN(time) ? 0 : time;
}

async function fetchCustomers() {
  if (!SHEET_DATA_URL) {
    console.warn('SHEET_DATA_URL is not set; returning empty customer list.');
    return [];
  }
  try {
    const response = await fetch(SHEET_DATA_URL, { headers: { Accept: 'application/json' } });
    if (!response.ok) return [];
    const payload = await response.json();
    const records = Array.isArray(payload?.records) ? payload.records : Array.isArray(payload) ? payload : [];
    return records.map((raw, index) => normaliseRecord(raw, index)).filter(record => record.customerName || record.licensePlate);
  } catch (error) {
    console.error('Error fetching customers from sheet:', error);
    return [];
  }
}

function normalisePath(pathname) {
  if (!pathname || pathname === '/') return '/';
  const trimmed = pathname.replace(/\/+$/, '');
  return trimmed === '' ? '/' : trimmed;
}

function validateFormData(parsed) {
  const formData = {
    customerName: parsed.customerName?.trim() ?? '',
    licensePlate: parsed.licensePlate?.trim() ?? '',
    actIssuedDate: parsed.actIssuedDate?.trim() ?? '',
    actExpiryDate: parsed.actExpiryDate?.trim() ?? '',
    taxRenewalDate: parsed.taxRenewalDate?.trim() ?? '',
    taxExpiryDate: parsed.taxExpiryDate?.trim() ?? '',
    voluntaryIssuedDate: parsed.voluntaryIssuedDate?.trim() ?? '',
    voluntaryExpiryDate: parsed.voluntaryExpiryDate?.trim() ?? '',
    phone: parsed.phone?.trim() ?? '',
    notes: parsed.notes?.trim() ?? '',
    rowNumber: parsed.rowNumber ? Number.parseInt(parsed.rowNumber, 10) : undefined,
    timestamp: parsed.timestamp?.trim() ?? ''
  };

  const errors = {};

  if (!formData.customerName) errors.customerName = 'กรุณากรอกชื่อลูกค้า';
  if (!formData.licensePlate) errors.licensePlate = 'กรุณากรอกทะเบียนรถ';

  const normaliseDateField = (field, label, required = false) => {
    const formatted = toInputDate(formData[field]);
    formData[`${field}Input`] = formatted;
    if (!formatted) {
      if (required) errors[field] = `กรุณาเลือก${label}`;
      formData[field] = '';
      return '';
    }
    formData[field] = formatted;
    return formatted;
  };

  const actIssued = normaliseDateField('actIssuedDate', 'วันที่ทำ พ.ร.บ.');
  const actExpiry = normaliseDateField('actExpiryDate', 'วันที่ครบกำหนด พ.ร.บ.');
  const taxRenewal = normaliseDateField('taxRenewalDate', 'วันที่ต่อภาษี');
  const taxExpiry = normaliseDateField('taxExpiryDate', 'วันที่ครบกำหนดต่อภาษี');
  const voluntaryIssued = normaliseDateField('voluntaryIssuedDate', 'วันที่ทำกรมธรรม์ภาคสมัครใจ');
  const voluntaryExpiry = normaliseDateField('voluntaryExpiryDate', 'วันที่ครบกำหนดกรมธรรม์ภาคสมัครใจ');

  if (!formData.phone) {
    errors.phone = 'กรุณากรอกเบอร์ติดต่อหลัก';
  } else {
    const digits = formData.phone.replace(/[^0-9]/g, '');
    if (digits.length < 7) errors.phone = 'กรุณากรอกเบอร์โทรศัพท์ที่ถูกต้อง';
  }

  return {
    formData: {
      ...formData,
      actIssuedDate: actIssued,
      actExpiryDate: actExpiry,
      taxRenewalDate: taxRenewal,
      taxExpiryDate: taxExpiry,
      voluntaryIssuedDate: voluntaryIssued,
      voluntaryExpiryDate: voluntaryExpiry
    },
    errors
  };
}

function handleCreateCustomer(req, res) {
  let body = '';
  req.on('data', chunk => {
    body += chunk.toString();
    if (body.length > 1e6) req.socket.destroy();
  });
  req.on('end', async () => {
    const parsed = parse(body);
    const { formData, errors } = validateFormData(parsed);
    if (Object.keys(errors).length > 0) {
      res.writeHead(400, { 'Content-Type': 'text/html; charset=utf-8' });
      res.end(renderAddCustomerPage({ message: 'กรุณาตรวจสอบข้อมูลที่ไฮไลต์และลองอีกครั้ง', status: 'error', formData, errors }));
      return;
    }
    const record = {
      timestamp: new Date().toISOString(),
      customerName: formData.customerName,
      licensePlate: formData.licensePlate,
      actIssuedDate: formData.actIssuedDate || null,
      actExpiryDate: formData.actExpiryDate || null,
      taxRenewalDate: formData.taxRenewalDate || null,
      taxExpiryDate: formData.taxExpiryDate || null,
      voluntaryIssuedDate: formData.voluntaryIssuedDate || null,
      voluntaryExpiryDate: formData.voluntaryExpiryDate || null,
      phone: formData.phone,
      notes: formData.notes || null
    };
    await syncToGoogleSheet(record);
    res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
    res.end(renderAddCustomerPage({ message: 'บันทึกข้อมูลลูกค้าเรียบร้อยแล้ว', status: 'success' }));
  });
}

function handleUpdateCustomer(req, res) {
  let body = '';
  req.on('data', chunk => {
    body += chunk.toString();
    if (body.length > 1e6) req.socket.destroy();
  });
  req.on('end', async () => {
    const parsed = parse(body);
    const { formData, errors } = validateFormData(parsed);
    if (!formData.rowNumber) errors.rowNumber = 'ไม่พบหมายเลขแถวของข้อมูล';

    if (Object.keys(errors).length > 0) {
      res.writeHead(400, { 'Content-Type': 'text/html; charset=utf-8' });
      res.end(renderEditCustomerPage({ message: 'กรุณาตรวจสอบข้อมูลที่ไฮไลต์และลองอีกครั้ง', status: 'error', formData, errors }));
      return;
    }
    const record = {
      timestamp: formData.timestamp || new Date().toISOString(),
      customerName: formData.customerName,
      licensePlate: formData.licensePlate,
      actIssuedDate: formData.actIssuedDate || null,
      actExpiryDate: formData.actExpiryDate || null,
      taxRenewalDate: formData.taxRenewalDate || null,
      taxExpiryDate: formData.taxExpiryDate || null,
      voluntaryIssuedDate: formData.voluntaryIssuedDate || null,
      voluntaryExpiryDate: formData.voluntaryExpiryDate || null,
      phone: formData.phone,
      notes: formData.notes || null
    };
    await updateCustomerInSheet(formData.rowNumber, record);
    res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
    res.end(renderEditCustomerPage({ message: 'บันทึกการแก้ไขเรียบร้อยแล้ว', status: 'success', formData }));
  });
}

function daysUntil(dateIso) {
  if (!dateIso) return null;
  const target = new Date(dateIso);
  if (Number.isNaN(target.getTime())) return null;
  const today = new Date();
  const diffMs = target.setHours(0,0,0,0) - today.setHours(0,0,0,0);
  return Math.ceil(diffMs / (1000 * 60 * 60 * 24));
}

async function handleExpiring(req, res, url) {
  const urlParams = new URLSearchParams(url.search || '');
  const days = Math.min(365, Math.max(1, Number.parseInt(urlParams.get('days') || String(DEFAULT_EXPIRY_WINDOW_DAYS), 10)));
  const customers = await fetchCustomers();
  const items = customers
    .map(c => ({ customer: c, daysRemaining: daysUntil(c.actExpiryDate) }))
    .filter(x => x.daysRemaining !== null && x.daysRemaining >= 0 && x.daysRemaining <= days)
    .sort((a, b) => (a.daysRemaining - b.daysRemaining) || (getRecordSortTime(a.customer) - getRecordSortTime(b.customer)));
  const html = renderExpiringPage({ customers: items, days });
  res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
  res.end(html);
}

async function handleSearch(req, res, url) {
  const q = (new URLSearchParams(url.search || '')).get('q')?.trim() || '';
  const customers = await fetchCustomers();
  const results = q
    ? customers.filter(r =>
        String(r.customerName || '').toLowerCase().includes(q.toLowerCase()) ||
        String(r.licensePlate || '').toLowerCase().includes(q.toLowerCase())
      )
    : customers;
  const html = renderSearchPage({ query: q, results, total: customers.length });
  res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
  res.end(html);
}
const server = createServer(async (req, res) => {
  try {
    const url = new URL(req.url, `http://${req.headers.host}`);
    const path = normalisePath(url.pathname);

    if (req.method === 'GET' && path === '/') {
      res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
      res.end(renderLayout({ pageTitle: 'หน้าหลัก', active: 'home', content: renderHomePage().match(/<article class=\"card\">([\s\S]*)<\/article>/)?.[1] || '' }));
      return;
    }

    if (req.method === 'GET' && path === '/customers/new') {
      res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
      res.end(renderAddCustomerPage());
      return;
    }

    if (req.method === 'POST' && path === '/customers') return handleCreateCustomer(req, res);

    if (req.method === 'GET' && path === '/customers/edit') {
      const row = Number.parseInt(new URLSearchParams(url.search || '').get('row') || '', 10);
      const customers = await fetchCustomers();
      const found = customers.find(c => Number(c.rowNumber) === row);
      if (!found) {
        res.writeHead(404, { 'Content-Type': 'text/html; charset=utf-8' });
        res.end(renderLayout({ pageTitle: 'ไม่พบข้อมูล', active: 'search', content: '<p class="muted">ไม่พบลูกค้าที่ต้องการแก้ไข</p>' }));
        return;
      }
      res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
      res.end(renderEditCustomerPage({ formData: found }));
      return;
    }

    if (req.method === 'POST' && path === '/customers/update') return handleUpdateCustomer(req, res);

    if (req.method === 'GET' && path === '/customers/search') return handleSearch(req, res, url);

    if (req.method === 'GET' && path === '/customers/expiring') return handleExpiring(req, res, url);

    res.writeHead(404, { 'Content-Type': 'text/html; charset=utf-8' });
    res.end(renderLayout({ pageTitle: 'ไม่พบหน้า', content: '<p class="muted">ไม่พบหน้า</p>' }));
  } catch (err) {
    console.error('Request error:', err);
    res.writeHead(500, { 'Content-Type': 'text/html; charset=utf-8' });
    res.end(renderLayout({ pageTitle: 'เกิดข้อผิดพลาด', content: '<p class="muted">เกิดข้อผิดพลาดภายในเซิร์ฟเวอร์</p>' }));
  }
});

server.listen(PORT, HOST, () => {
  console.log(`Server running at http://${HOST}:${PORT}`);
});
