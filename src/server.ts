/**
 * Monchai Insurance - TypeScript Server
 * 
 * This is a refactored TypeScript version with improved performance:
 * - Modular architecture (separate files for types, utils, services, handlers)
 * - LRU caching for customer data (30s TTL) - reduces API calls significantly
 * - Strong typing for better code quality and maintainability
 * - Proper error handling
 */

import { createServer, IncomingMessage, ServerResponse } from 'http';
import { URL } from 'url';
import { config } from './config.js';
import { normalisePath, daysUntil } from './utils/helpers.js';
import { fetchCustomers } from './services/customerService.js';
import { runBackupOnStartup, saveCustomersToFile } from './services/backupService.js';
import { renderLayout, renderHomePage } from './views/layout.js';
import {
  handleCreateCustomer,
  handleUpdateCustomer,
  handleDeleteCustomers,
} from './handlers/customerHandlers.js';
import { Customer } from './types/index.js';

/**
 * Renders the search page with customer list
 */
function renderSearchPage(query: string, results: Customer[], total: number): string {
  const tableContent = results.length
    ? `<table class="table">
        <thead>
          <tr>
            <th><input type="checkbox" id="selectAll" /></th>
            <th>ชื่อลูกค้า</th>
            <th>ทะเบียนรถ</th>
            <th>เบอร์โทร</th>
            <th>หมายเหตุ</th>
          </tr>
        </thead>
        <tbody>
          ${results.map(r => `
            <tr>
              <td><input type="checkbox" name="rows" value="${r.rowNumber || ''}" form="bulkDeleteForm" /></td>
              <td><a class="record-link" href="/customers/edit?row=${r.rowNumber || ''}">${r.customerName || ''}</a></td>
              <td>${r.licensePlate || ''}</td>
              <td>${r.phone || ''}</td>
              <td>${r.notes || ''}</td>
            </tr>`).join('')}
        </tbody>
      </table>
      <form id="bulkDeleteForm" method="POST" action="/customers/delete">
        <button type="submit" class="primary">ลบที่เลือก</button>
      </form>`
    : `<p class="muted">ไม่พบข้อมูลลูกค้า</p>`;

  const content = `
    <h2>ค้นหาลูกค้า</h2>
    <form method="GET" action="/customers/search">
      <input type="text" name="q" value="${query}" placeholder="ค้นหาชื่อหรือทะเบียนรถ" />
      <button type="submit" class="primary">ค้นหา</button>
    </form>
    <p class="muted">พบ ${results.length} จาก ${total} รายการ</p>
    ${tableContent}
  `;
  return renderLayout({ pageTitle: 'ค้นหาลูกค้า', active: 'search', content });
}

/**
 * Renders the expiring customers page
 */
function renderExpiringPage(customers: Array<{customer: Customer, minDaysRemaining: number | null, act: number | null, tax: number | null, vol: number | null}>): string {
  const tableContent = customers.length
    ? `<table class="table">
        <thead>
          <tr>
            <th>ชื่อลูกค้า</th>
            <th>ทะเบียนรถ</th>
            <th>เหลืออีก (วัน)</th>
            <th>เบอร์ติดต่อ</th>
          </tr>
        </thead>
        <tbody>
          ${customers.map(item => `
            <tr>
              <td><a class="record-link" href="/customers/edit?row=${item.customer.rowNumber || ''}&from=expiring">${item.customer.customerName || ''}</a></td>
              <td>${item.customer.licensePlate || ''}</td>
              <td>${item.minDaysRemaining == null ? '-' : item.minDaysRemaining}</td>
              <td>${item.customer.phone || ''}</td>
            </tr>`).join('')}
        </tbody>
      </table>`
    : `<p class="muted">ไม่มีลูกค้าที่จะหมดอายุภายใน ${config.DEFAULT_EXPIRY_WINDOW_DAYS} วัน</p>`;

  const content = `
    <h2>แจ้งเตือนลูกค้าที่จะหมดอายุ</h2>
    <p class="lead">รายการที่ครบกำหนดภายใน ${config.DEFAULT_EXPIRY_WINDOW_DAYS} วันข้างหน้า</p>
    ${tableContent}
  `;
  return renderLayout({ pageTitle: 'แจ้งเตือนลูกค้าที่จะหมดอายุ', active: 'expiring', content });
}

/**
 * Handle search requests
 */
async function handleSearch(_req: IncomingMessage, res: ServerResponse, url: URL): Promise<void> {
  const q = url.searchParams.get('q')?.trim() || '';
  const customers = await fetchCustomers();
  const results = q
    ? customers.filter(r =>
        String(r.customerName || '').toLowerCase().includes(q.toLowerCase()) ||
        String(r.licensePlate || '').toLowerCase().includes(q.toLowerCase())
      )
    : customers;

  res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
  res.end(renderSearchPage(q, results, customers.length));
}

/**
 * Handle expiring customers view
 */
async function handleExpiring(_req: IncomingMessage, res: ServerResponse): Promise<void> {
  const customers = await fetchCustomers();
  const expiring = customers
    .map(customer => {
      const act = daysUntil(customer.actExpiryDate);
      const tax = daysUntil(customer.taxExpiryDate);
      const vol = daysUntil(customer.voluntaryExpiryDate);
      
      const validDays = [act, tax, vol].filter(d => d !== null) as number[];
      const minDaysRemaining = validDays.length > 0 ? Math.min(...validDays) : null;
      
      return { customer, minDaysRemaining, act, tax, vol };
    })
    .filter(item => 
      item.minDaysRemaining !== null && 
      item.minDaysRemaining >= 0 && 
      item.minDaysRemaining < config.DEFAULT_EXPIRY_WINDOW_DAYS
    )
    .sort((a, b) => (a.minDaysRemaining || 0) - (b.minDaysRemaining || 0));

  res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
  res.end(renderExpiringPage(expiring));
}

/**
 * Handle backup endpoint
 */
async function handleBackup(_req: IncomingMessage, res: ServerResponse): Promise<void> {
  try {
    const customers = await fetchCustomers();
    const filepath = await saveCustomersToFile(customers);
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ success: true, filepath, count: customers.length }));
  } catch (error) {
    console.error('Backup error:', error);
    res.writeHead(500, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ success: false, error: String(error) }));
  }
}

/**
 * Handle edit customer page
 */
async function handleEditCustomer(_req: IncomingMessage, res: ServerResponse, url: URL): Promise<void> {
  const params = url.searchParams;
  const row = Number.parseInt(params.get('row') || '', 10);
  const from = params.get('from') || '';
  
  const customers = await fetchCustomers();
  const found = customers.find(c => c.rowNumber === row);
  
  if (!found) {
    res.writeHead(404, { 'Content-Type': 'text/html; charset=utf-8' });
    res.end(renderLayout({ 
      pageTitle: 'ไม่พบข้อมูล', 
      active: 'search', 
      content: '<p class="muted">ไม่พบลูกค้าที่ต้องการแก้ไข</p>' 
    }));
    return;
  }

  // For now, show a simple edit form placeholder
  const content = `
    <h2>แก้ไขข้อมูลลูกค้า</h2>
    <p>Customer: ${found.customerName}</p>
    <p>License Plate: ${found.licensePlate}</p>
    <a href="/customers/search">กลับไปค้นหา</a>
  `;
  
  res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
  res.end(renderLayout({ 
    pageTitle: 'แก้ไขข้อมูล', 
    active: from === 'expiring' ? 'expiring' : 'search', 
    content 
  }));
}

/**
 * Main HTTP server
 */
const server = createServer(async (req: IncomingMessage, res: ServerResponse) => {
  try {
    const url = new URL(req.url || '/', `http://${req.headers.host}`);
    const path = normalisePath(url.pathname);

    // Routes
    if (req.method === 'GET' && path === '/') {
      res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
      res.end(renderHomePage());
      return;
    }

    if (req.method === 'GET' && path === '/customers/new') {
      res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
      res.end(renderLayout({ 
        pageTitle: 'เพิ่มลูกค้าใหม่', 
        active: 'new', 
        content: '<h2>เพิ่มลูกค้าใหม่</h2><p>Form coming soon...</p>' 
      }));
      return;
    }

    if (req.method === 'POST' && path === '/customers') {
      return handleCreateCustomer(req, res);
    }

    if (req.method === 'GET' && path === '/customers/edit') {
      return handleEditCustomer(req, res, url);
    }

    if (req.method === 'POST' && path === '/customers/update') {
      return handleUpdateCustomer(req, res);
    }

    if (req.method === 'GET' && path === '/customers/search') {
      return handleSearch(req, res, url);
    }

    if (req.method === 'POST' && path === '/customers/delete') {
      return handleDeleteCustomers(req, res);
    }

    if (req.method === 'GET' && path === '/customers/expiring') {
      return handleExpiring(req, res);
    }

    if (req.method === 'GET' && path === '/api/backup') {
      return handleBackup(req, res);
    }

    // 404
    res.writeHead(404, { 'Content-Type': 'text/html; charset=utf-8' });
    res.end(renderLayout({ 
      pageTitle: 'ไม่พบหน้า', 
      content: '<p class="muted">ไม่พบหน้า</p>' 
    }));
  } catch (err) {
    console.error('Request error:', err);
    res.writeHead(500, { 'Content-Type': 'text/html; charset=utf-8' });
    res.end(renderLayout({ 
      pageTitle: 'เกิดข้อผิดพลาด', 
      content: '<p class="muted">เกิดข้อผิดพลาดภายในเซิร์ฟเวอร์</p>' 
    }));
  }
});

// Start server
server.listen(config.PORT, config.HOST, async () => {
  console.log(`?? TypeScript Server running at http://${config.HOST}:${config.PORT}`);
  console.log(`?? Caching enabled (30s TTL) for better performance`);
  console.log(`?? Environment: ${process.env.NODE_ENV || 'development'}`);
  await runBackupOnStartup();
});

// Graceful shutdown
process.on('SIGTERM', () => {
  console.log('SIGTERM signal received: closing HTTP server');
  server.close(() => {
    console.log('HTTP server closed');
  });
});
