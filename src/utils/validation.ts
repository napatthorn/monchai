import { FormData, ValidationErrors } from '../types/index.js';
import { toInputDate } from './helpers.js';

/**
 * Validates form data and returns normalized form data with any errors
 */
export function validateFormData(parsed: Record<string, unknown>): {
  formData: FormData;
  errors: ValidationErrors;
} {
  const formData: FormData = {
    customerName: String(parsed.customerName || '').trim(),
    policyNumber: String(parsed.policyNumber || '').trim(),
    licensePlate: String(parsed.licensePlate || '').trim(),
    actIssuedDate: String(parsed.actIssuedDate || '').trim(),
    actExpiryDate: String(parsed.actExpiryDate || '').trim(),
    taxRenewalDate: String(parsed.taxRenewalDate || '').trim(),
    taxExpiryDate: String(parsed.taxExpiryDate || '').trim(),
    voluntaryIssuedDate: String(parsed.voluntaryIssuedDate || '').trim(),
    voluntaryExpiryDate: String(parsed.voluntaryExpiryDate || '').trim(),
    registrationDate: String(parsed.registrationDate || '').trim(),
    email: String(parsed.email || '').trim(),
    phone: String(parsed.phone || '').trim(),
    status: String(parsed.status || '').trim(),
    notes: String(parsed.notes || '').trim(),
  };

  if (parsed.rowNumber) formData.rowNumber = String(parsed.rowNumber);
  if (parsed.from) formData.from = String(parsed.from);
  if (parsed.timestamp) formData.timestamp = String(parsed.timestamp);

  // Normalize dates
  formData.actIssuedDate = toInputDate(formData.actIssuedDate);
  formData.actExpiryDate = toInputDate(formData.actExpiryDate);
  formData.taxRenewalDate = toInputDate(formData.taxRenewalDate);
  formData.taxExpiryDate = toInputDate(formData.taxExpiryDate);
  formData.voluntaryIssuedDate = toInputDate(formData.voluntaryIssuedDate);
  formData.voluntaryExpiryDate = toInputDate(formData.voluntaryExpiryDate);
  formData.registrationDate = toInputDate(formData.registrationDate);

  const errors: ValidationErrors = {};

  // Validation rules
  if (!formData.customerName && !formData.licensePlate) {
    errors.customerName = 'กรุณากรอกชื่อลูกค้าหรือทะเบียนรถอย่างน้อย 1 อย่าง';
    errors.licensePlate = 'กรุณากรอกทะเบียนรถหรือชื่อลูกค้าอย่างน้อย 1 อย่าง';
  }

  if (!formData.phone) {
    errors.phone = 'กรุณากรอกเบอร์โทร';
  }

  return { formData, errors };
}
