/**
 * Escapes HTML special characters to prevent XSS attacks
 */
export function escapeHtml(value: unknown = ''): string {
  const text = value == null ? '' : String(value);
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/**
 * Formats ISO date string to Thai locale format
 */
export function formatDate(dateIso: string | null | undefined): string {
  if (!dateIso) return '-';
  const date = new Date(dateIso);
  if (Number.isNaN(date.getTime())) return '-';
  return date.toLocaleDateString('th-TH', { 
    year: 'numeric', 
    month: 'short', 
    day: 'numeric' 
  });
}

/**
 * Converts various date formats to YYYY-MM-DD format for input fields
 */
export function toInputDate(value: string | Date | null | undefined): string {
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

/**
 * Adds one year to a date, keeping the same day
 */
export function addOneYearSameDay(dateIso: string | null | undefined): string {
  if (!dateIso) return '';
  const base = new Date(dateIso);
  if (Number.isNaN(base.getTime())) return '';
  
  const next = new Date(base);
  next.setFullYear(next.getFullYear() + 1);
  return toInputDate(next);
}

/**
 * Calculates days until a date from today
 */
export function daysUntil(dateIso: string | null | undefined): number | null {
  if (!dateIso) return null;
  const target = new Date(dateIso);
  if (Number.isNaN(target.getTime())) return null;
  
  const now = new Date();
  now.setHours(0, 0, 0, 0);
  target.setHours(0, 0, 0, 0);
  
  const diffMs = target.getTime() - now.getTime();
  return Math.floor(diffMs / (1000 * 60 * 60 * 24));
}

/**
 * Normalizes URL pathname by removing trailing slashes
 */
export function normalisePath(pathname: string | undefined): string {
  if (!pathname || pathname === '/') return '/';
  const trimmed = pathname.replace(/\/+$/, '');
  return trimmed === '' ? '/' : trimmed;
}
