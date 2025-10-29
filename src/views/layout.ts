import { config } from '../config.js';
import { escapeHtml } from '../utils/helpers.js';
import { RenderOptions } from '../types/index.js';

/**
 * Main layout template
 */
export function renderLayout({ pageTitle, active = '', content = '', alertMarkup = '' }: RenderOptions): string {
  const navItems = [
    { href: '/', label: 'หนาหลก', key: 'home' },
    { href: '/customers/new', label: 'เพมลกคาใหม', key: 'new' },
    { href: '/customers/search', label: 'คนหาลกคา', key: 'search' },
    { href: '/customers/expiring', label: 'แจงเตอนลกคาทจะหมดอาย', key: 'expiring' }
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
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/flatpickr/dist/flatpickr.min.css">
  <script src="https://cdn.jsdelivr.net/npm/flatpickr"></script>
  <script src="https://cdn.jsdelivr.net/npm/flatpickr/dist/l10n/th.js"></script>
  <link rel="icon" type="image/svg+xml" href="data:image/svg+xml,%3Csvg%20xmlns%3D%22http%3A%2F%2Fwww.w3.org%2F2000%2Fsvg%22%20viewBox%3D%220%200%2064%2064%22%3E%0A%20%20%3Cdefs%3E%0A%20%20%20%20%3ClinearGradient%20id%3D%22g%22%20x1%3D%220%22%20x2%3D%221%22%20y1%3D%220%22%20y2%3D%221%22%3E%0A%20%20%20%20%20%20%3Cstop%20offset%3D%220%25%22%20stop-color%3D%22%232563eb%22%2F%3E%0A%20%20%20%20%20%20%3Cstop%20offset%3D%22100%25%22%20stop-color%3D%22%2338bdf8%22%2F%3E%0A%20%20%20%20%3C%2FlinearGradient%3E%0A%20%20%3C%2Fdefs%3E%0A%20%20%3Crect%20width%3D%2264%22%20height%3D%2264%22%20rx%3D%2214%22%20fill%3D%22url%28%23g%29%22%2F%3E%0A%20%20%3Cpath%20d%3D%22M20%2036c6-1%2010-6%2012-12%202%206%206%2011%2012%2012-6%201-10%206-12%2012-2-6-6-11-12-12z%22%20fill%3D%22%23f8fafc%22%2F%3E%0A%3C%2Fsvg%3E" />
  <title>${escapeHtml(pageTitle)} - ${config.APP_TITLE}</title>
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
    .button.secondary { display:inline-flex; align-items:center; justify-content:center; padding:0.7rem 1.2rem; border-radius:10px; border:1px solid #94a3b8; background:#f1f5f9; color:#1e293b; text-decoration:none; font-size:0.95rem; transition:all 0.2s ease; margin-left: 0.5rem; }
    .button.secondary:hover { background:#e2e8f0; transform: translateY(-1px); box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    label { font-weight: 600; font-size: 0.95rem; color: #243b53; }
    input[type="text"], input[type="tel"], input[type="date"], input[type="email"], textarea, select { padding: 0.8rem 0.95rem; font-size: 1rem; border-radius: 10px; border: 1px solid #d9e2ec; background: #f8fafc; transition: border-color 0.2s ease, box-shadow 0.2s ease; resize: vertical; min-height: 48px; }
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
    .req { color: #e53e3e; margin-left: 4px; font-weight: 700; }
    .due-highlight { color: #e53e3e; font-weight: 700; }
    .btn-icon { display: inline-block; margin-right: 8px; }
    .icon-only .btn-text { display: none; }
    .icon-only .btn-icon { margin-right: 0; }
    .icon-only { padding: 0.5rem; width: 42px; min-width: 42px; display: inline-flex; align-items: center; justify-content: center; }
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
        <h1>${config.APP_TITLE}</h1>
        <div class="muted">วางแผนตดตอลกคาเพอการตออายกรมธรรม</div>
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
  <footer>Monchai Insurance  ระบบชวยเตอนตออายลกคา</footer>
</body>
</html>`;
}

/**
 * Home page view
 */
export function renderHomePage(): string {
  const content = `
    <h2>ภาพรวมการทำงาน</h2>
    <p class="lead">เลอกเมนทตองการจดการลกคา เรมจากบนทกลกคาใหม ไปจนถงคนหาและดรายการทใกลหมดอาย</p>
    <div class="menu-grid">
      <a class="menu-card" href="/customers/new">
        <div class="menu-icon"></div>
        <strong>เพมลกคาใหม</strong>
        <span>กรอกขอมลลกคาและตงบนทกการตดตอลวงหนา</span>
      </a>
      <a class="menu-card" href="/customers/search">
        <div class="menu-icon"></div>
        <strong>คนหาลกคา</strong>
        <span>ดขอมลลกคาทงหมดและคนหาตามชอหรอทะเบยนรถ</span>
      </a>
      <a class="menu-card" href="/customers/expiring">
        <div class="menu-icon"></div>
        <strong>แจงเตอนลกคาทจะหมดอาย</strong>
        <span>ดรายชอลกคาทมกำหนดตออายภายใน ${config.DEFAULT_EXPIRY_WINDOW_DAYS} วน</span>
      </a>
    </div>
  `;
  return renderLayout({ pageTitle: 'หนาหลก', active: 'home', content });
}