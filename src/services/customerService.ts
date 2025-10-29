import fetch from 'node-fetch';
import { Customer, SheetResponse } from '../types/index.js';
import { config } from '../config.js';
import { customerCache } from './cache.js';
import { toInputDate } from '../utils/helpers.js';

/**
 * Normalizes a raw record from Google Sheets to Customer format
 */
function normaliseRecord(raw: Record<string, unknown> = {}, index: number = 0): Customer {
  const rowNumber = Number.isFinite(raw.rowNumber) 
    ? Number(raw.rowNumber) 
    : index + 2;

  return {
    rowNumber,
    timestamp: String(raw.timestamp || new Date().toISOString()),
    customerName: String(raw.customerName || ''),
    licensePlate: String(raw.licensePlate || ''),
    policyNumber: raw.policyNumber ? String(raw.policyNumber) : null,
    actIssuedDate: raw.actIssuedDate ? toInputDate(String(raw.actIssuedDate)) : null,
    actExpiryDate: raw.actExpiryDate ? toInputDate(String(raw.actExpiryDate)) : null,
    taxRenewalDate: raw.taxRenewalDate ? toInputDate(String(raw.taxRenewalDate)) : null,
    taxExpiryDate: raw.taxExpiryDate ? toInputDate(String(raw.taxExpiryDate)) : null,
    voluntaryIssuedDate: raw.voluntaryIssuedDate ? toInputDate(String(raw.voluntaryIssuedDate)) : null,
    voluntaryExpiryDate: raw.voluntaryExpiryDate ? toInputDate(String(raw.voluntaryExpiryDate)) : null,
    registrationDate: raw.registrationDate ? toInputDate(String(raw.registrationDate)) : null,
    email: raw.email ? String(raw.email) : null,
    phone: String(raw.phone || ''),
    status: raw.status ? String(raw.status) : null,
    notes: raw.notes ? String(raw.notes) : null,
  };
}

/**
 * Gets sort time for a customer record (used for ordering)
 */
function getRecordSortTime(record: Customer): number {
  const ts = record.timestamp ? new Date(record.timestamp).getTime() : 0;
  return Number.isNaN(ts) ? 0 : ts;
}

/**
 * Fetches all customers from Google Sheets with caching
 */
export async function fetchCustomers(): Promise<Customer[]> {
  // Check cache first for performance
  const cached = customerCache.get();
  if (cached) {
    return cached;
  }

  if (!config.SHEET_DATA_URL) {
    console.warn('SHEET_DATA_URL is not set; returning empty customer list.');
    return [];
  }

  try {
    const response = await fetch(config.SHEET_DATA_URL, { 
      headers: { Accept: 'application/json' } 
    });
    
    if (!response.ok) return [];
    
    const payload = await response.json() as SheetResponse | Customer[];
    const records = Array.isArray((payload as SheetResponse)?.records) 
      ? (payload as SheetResponse).records! 
      : Array.isArray(payload) 
        ? payload 
        : [];

    const customers = records
      .map((raw, index) => normaliseRecord(raw as unknown as Record<string, unknown>, index))
      .filter(record => record.customerName || record.licensePlate)
      .sort((a, b) => getRecordSortTime(b) - getRecordSortTime(a));

    // Cache the results
    customerCache.set(customers);
    
    return customers;
  } catch (error) {
    console.error('Error fetching customers from sheet:', error);
    return [];
  }
}

/**
 * Syncs a new customer record to Google Sheets
 */
export async function syncToGoogleSheet(record: Customer): Promise<void> {
  if (!config.SHEET_WRITE_URL) {
    console.warn('SHEET_WRITE_URL is not set; skipping sync.');
    return;
  }

  await fetch(config.SHEET_WRITE_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(record),
  });

  // Invalidate cache after write
  customerCache.invalidate();
}

/**
 * Updates an existing customer in Google Sheets
 */
export async function updateCustomerInSheet(rowNumber: number, record: Customer): Promise<void> {
  if (!config.SHEET_UPDATE_URL) {
    console.warn('SHEET_UPDATE_URL is not set; skipping update.');
    return;
  }

  await fetch(config.SHEET_UPDATE_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ ...record, rowNumber }),
  });

  // Invalidate cache after update
  customerCache.invalidate();
}

/**
 * Deletes multiple customers from Google Sheets
 */
export async function deleteCustomersInSheet(rowNumbers: number[] = []): Promise<void> {
  if (!config.SHEET_WRITE_URL || rowNumbers.length === 0) return;

  await fetch(config.SHEET_WRITE_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ action: 'delete', rows: rowNumbers }),
  });

  // Invalidate cache after delete
  customerCache.invalidate();
}
