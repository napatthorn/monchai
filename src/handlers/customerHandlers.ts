import { IncomingMessage, ServerResponse } from 'http';
import { parse } from 'querystring';
import { config } from '../config.js';
import { Customer } from '../types/index.js';
import { validateFormData } from '../utils/validation.js';
import { daysUntil, escapeHtml } from '../utils/helpers.js';
import {
  fetchCustomers,
  syncToGoogleSheet,
  updateCustomerInSheet,
  deleteCustomersInSheet,
} from '../services/customerService.js';
import { renderAddCustomerPage, renderEditCustomerPage } from '../views/customerViews.js';

/**
 * Parse request body from POST requests
 */
async function parseBody(req: IncomingMessage): Promise<Record<string, unknown>> {
  return new Promise((resolve, reject) => {
    let body = '';
    req.on('data', (chunk) => {
      body += chunk.toString();
      if (body.length > 1e6) {
        req.socket?.destroy();
        reject(new Error('Request too large'));
      }
    });
    req.on('end', () => {
      try {
        resolve(parse(body) as Record<string, unknown>);
      } catch (error) {
        reject(error);
      }
    });
    req.on('error', reject);
  });
}

/**
 * Handles creating a new customer
 */
export async function handleCreateCustomer(
  req: IncomingMessage,
  res: ServerResponse
): Promise<void> {
  try {
    const parsed = await parseBody(req);
    const { formData, errors } = validateFormData(parsed);

    if (Object.keys(errors).length > 0) {
      res.writeHead(400, { 'Content-Type': 'text/html; charset=utf-8' });
      res.end(
        renderAddCustomerPage({
          message: 'กรุณาตรวจสอบข้อมูลที่ไฮไลต์และลองอีกครั้ง',
          status: 'error',
          formData,
          errors,
        })
      );
      return;
    }

    // Check for duplicate license plate
    if (formData.licensePlate) {
      const existingCustomers = await fetchCustomers();
      const duplicate = existingCustomers.find(
        (customer) =>
          customer.licensePlate &&
          customer.licensePlate.toLowerCase() === formData.licensePlate.toLowerCase()
      );

      if (duplicate) {
        errors.licensePlate = 'ทะเบียนรถนี้มีอยู่ในระบบแล้ว';
        res.writeHead(400, { 'Content-Type': 'text/html; charset=utf-8' });
        res.end(
          renderAddCustomerPage({
            message: `ทะเบียนรถ "${escapeHtml(formData.licensePlate)}" มีอยู่ในระบบแล้ว (ลูกค้า: ${escapeHtml(
              duplicate.customerName || 'ไม่ระบุชื่อ'
            )})`,
            status: 'error',
            formData,
            errors,
          })
        );
        return;
      }
    }

    // Business rule validation
    if ((formData.status || '') === 'ต่อสัญญาเรียบร้อย') {
      const act = daysUntil(formData.actExpiryDate);
      const tax = daysUntil(formData.taxExpiryDate);
      const vol = daysUntil(formData.voluntaryExpiryDate);
      const within = (v: number | null) =>
        v !== null && v >= 0 && v < config.DEFAULT_EXPIRY_WINDOW_DAYS;

      if (within(act) || within(tax) || within(vol)) {
        const errMsg = `ปรับวันที่ให้พ้นช่วงแจ้งเตือน (< ${config.DEFAULT_EXPIRY_WINDOW_DAYS} วัน) หรือเปลี่ยนสถานะ`;
        if (within(act)) errors.actExpiryDate = errMsg;
        if (within(tax)) errors.taxExpiryDate = errMsg;
        if (within(vol)) errors.voluntaryExpiryDate = errMsg;

        res.writeHead(400, { 'Content-Type': 'text/html; charset=utf-8' });
        res.end(
          renderAddCustomerPage({
            message: `ไม่สามารถบันทึกสถานะ "${escapeHtml(
              formData.status
            )}" ได้ เนื่องจากยังอยู่ในช่วงแจ้งเตือน (< ${
              config.DEFAULT_EXPIRY_WINDOW_DAYS
            } วัน) ของ พ.ร.บ./ต่อภาษี/ภาคสมัครใจ`,
            status: 'error',
            formData,
            errors,
          })
        );
        return;
      }
    }

    const record: Customer = {
      timestamp: new Date().toISOString(),
      customerName: formData.customerName,
      licensePlate: formData.licensePlate,
      policyNumber: formData.policyNumber || null,
      actIssuedDate: formData.actIssuedDate || null,
      actExpiryDate: formData.actExpiryDate || null,
      taxRenewalDate: formData.taxRenewalDate || null,
      taxExpiryDate: formData.taxExpiryDate || null,
      voluntaryIssuedDate: formData.voluntaryIssuedDate || null,
      voluntaryExpiryDate: formData.voluntaryExpiryDate || null,
      registrationDate: formData.registrationDate || null,
      email: formData.email || null,
      phone: formData.phone,
      status: formData.status || 'ยังไม่แจ้งลูกค้า',
      notes: formData.notes || null,
    };

    await syncToGoogleSheet(record);

    res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
    res.end(
      renderAddCustomerPage({
        message: 'บันทึกข้อมูลลูกค้าเรียบร้อยแล้ว',
        status: 'success',
      })
    );
  } catch (error) {
    console.error('Error creating customer:', error);
    res.writeHead(500, { 'Content-Type': 'text/html; charset=utf-8' });
    res.end(
      renderAddCustomerPage({
        message: 'เกิดข้อผิดพลาดในการบันทึกข้อมูล',
        status: 'error',
      })
    );
  }
}

/**
 * Handles updating an existing customer
 */
export async function handleUpdateCustomer(
  req: IncomingMessage,
  res: ServerResponse
): Promise<void> {
  try {
    const parsed = await parseBody(req);
    const { formData, errors } = validateFormData(parsed);
    const cameFromExpiring =
      String(formData.from || '')
        .trim()
        .toLowerCase() === 'expiring';

    // Handle "ลูกค้าไม่ต่อ" status logic
    if (!cameFromExpiring) {
      const statusTrimmed = String(formData.status || '').trim();
      if (statusTrimmed === 'ลูกค้าไม่ต่อ') {
        const suffix = ' (ลูกค้าไม่ต่อ)';
        const act = daysUntil(formData.actExpiryDate);
        const tax = daysUntil(formData.taxExpiryDate);
        const vol = daysUntil(formData.voluntaryExpiryDate);
        const within = (v: number | null) =>
          v !== null && v >= 0 && v < config.DEFAULT_EXPIRY_WINDOW_DAYS;

        if (within(act) || within(tax) || within(vol)) {
          const warning = `* ปรับวันที่ให้พ้นช่วงแจ้งเตือน (< ${config.DEFAULT_EXPIRY_WINDOW_DAYS} วัน)`;
          if (within(act)) errors.actExpiryDate = warning;
          if (within(tax)) errors.taxExpiryDate = warning;
          if (within(vol)) errors.voluntaryExpiryDate = warning;

          res.writeHead(400, { 'Content-Type': 'text/html; charset=utf-8' });
          res.end(
            renderEditCustomerPage({
              message: `ไม่สามารถบันทึกได้ เนื่องจากวันที่ครบกำหนดยังอยู่ในช่วงแจ้งเตือน (< ${config.DEFAULT_EXPIRY_WINDOW_DAYS} วัน)`,
              status: 'error',
              formData,
              errors,
              showStatus: formData.from === 'expiring',
              activeNav: formData.from === 'expiring' ? 'expiring' : 'search',
              from: formData.from || '',
            })
          );
          return;
        }

        const nameRaw = formData.customerName || '';
        const trimmedName = nameRaw.trimEnd();
        formData.customerName = trimmedName.endsWith(suffix)
          ? trimmedName.slice(0, -suffix.length).trimEnd()
          : trimmedName;
        formData.status = '';
      } else {
        formData.status = '';
      }
    }

    if (!formData.rowNumber) errors.rowNumber = 'ไม่พบหมายเลขแถวของข้อมูล';

    if (Object.keys(errors).length > 0) {
      res.writeHead(400, { 'Content-Type': 'text/html; charset=utf-8' });
      res.end(
        renderEditCustomerPage({
          message: 'กรุณาตรวจสอบข้อมูลที่ไฮไลต์และลองอีกครั้ง',
          status: 'error',
          formData,
          errors,
          showStatus: formData.from === 'expiring',
          activeNav: formData.from === 'expiring' ? 'expiring' : 'search',
          from: formData.from || '',
        })
      );
      return;
    }

    // Check for duplicate license plate (excluding current customer)
    if (formData.licensePlate) {
      const existingCustomers = await fetchCustomers();
      const duplicate = existingCustomers.find(
        (customer) =>
          customer.licensePlate &&
          customer.licensePlate.toLowerCase() === formData.licensePlate.toLowerCase() &&
          customer.rowNumber !== Number(formData.rowNumber)
      );

      if (duplicate) {
        errors.licensePlate = 'ทะเบียนรถนี้มีอยู่ในระบบแล้ว';
        res.writeHead(400, { 'Content-Type': 'text/html; charset=utf-8' });
        res.end(
          renderEditCustomerPage({
            message: `ทะเบียนรถ "${escapeHtml(formData.licensePlate)}" มีอยู่ในระบบแล้ว (ลูกค้า: ${escapeHtml(
              duplicate.customerName || 'ไม่ระบุชื่อ'
            )})`,
            status: 'error',
            formData,
            errors,
            showStatus: formData.from === 'expiring',
            activeNav: formData.from === 'expiring' ? 'expiring' : 'search',
            from: formData.from || '',
          })
        );
        return;
      }
    }

    // Business rule validation for expiring dashboard
    if (cameFromExpiring && (formData.status || '') === 'ต่อสัญญาเรียบร้อย') {
      const act = daysUntil(formData.actExpiryDate);
      const tax = daysUntil(formData.taxExpiryDate);
      const vol = daysUntil(formData.voluntaryExpiryDate);
      const within = (v: number | null) =>
        v !== null && v >= 0 && v < config.DEFAULT_EXPIRY_WINDOW_DAYS;

      if (within(act) || within(tax) || within(vol)) {
        const errMsg = `ปรับวันที่ให้อยู่พ้นช่วงแจ้งเตือน (< ${config.DEFAULT_EXPIRY_WINDOW_DAYS} วัน) หรือเปลี่ยนสถานะ`;
        if (within(act)) errors.actExpiryDate = errMsg;
        if (within(tax)) errors.taxExpiryDate = errMsg;
        if (within(vol)) errors.voluntaryExpiryDate = errMsg;

        res.writeHead(400, { 'Content-Type': 'text/html; charset=utf-8' });
        res.end(
          renderEditCustomerPage({
            message: `ไม่สามารถบันทึกสถานะ "${escapeHtml(
              formData.status
            )}" ได้ เนื่องจากยังอยู่ในช่วงแจ้งเตือน (< ${
              config.DEFAULT_EXPIRY_WINDOW_DAYS
            } วัน) ของ พ.ร.บ./ต่อภาษี/ภาคสมัครใจ`,
            status: 'error',
            formData,
            errors,
            showStatus: formData.from === 'expiring',
            activeNav: formData.from === 'expiring' ? 'expiring' : 'search',
            from: formData.from || '',
          })
        );
        return;
      }

      formData.status = '';
    }

    // Handle suffix for "ลูกค้าไม่ต่อ"
    const suffixNotRenew = ' (ลูกค้าไม่ต่อ)';
    const currentNameRaw = formData.customerName || '';
    const currentNameTrimmed = currentNameRaw.trimEnd();

    if (String(formData.status || '').trim() === 'ลูกค้าไม่ต่อ') {
      if (!currentNameTrimmed.endsWith(suffixNotRenew)) {
        formData.customerName = currentNameTrimmed
          ? `${currentNameTrimmed}${suffixNotRenew}`
          : 'ลูกค้าไม่ต่อ';
      }
    } else if (currentNameTrimmed.endsWith(suffixNotRenew)) {
      formData.customerName = currentNameTrimmed
        .slice(0, -suffixNotRenew.length)
        .trimEnd();
    }

    const statusForRecord =
      formData.status === '' ? null : formData.status || 'ยังไม่แจ้งลูกค้า';

    const record: Customer = {
      timestamp: formData.timestamp || new Date().toISOString(),
      customerName: formData.customerName,
      licensePlate: formData.licensePlate,
      policyNumber: formData.policyNumber || null,
      actIssuedDate: formData.actIssuedDate || null,
      actExpiryDate: formData.actExpiryDate || null,
      taxRenewalDate: formData.taxRenewalDate || null,
      taxExpiryDate: formData.taxExpiryDate || null,
      voluntaryIssuedDate: formData.voluntaryIssuedDate || null,
      voluntaryExpiryDate: formData.voluntaryExpiryDate || null,
      registrationDate: formData.registrationDate || null,
      email: formData.email || null,
      phone: formData.phone,
      status: statusForRecord,
      notes: formData.notes || null,
    };

    await updateCustomerInSheet(Number(formData.rowNumber), record);

    res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
    res.end(
      renderEditCustomerPage({
        message: 'บันทึกการแก้ไขเรียบร้อยแล้ว',
        status: 'success',
        formData,
        showStatus: formData.from === 'expiring',
        activeNav: formData.from === 'expiring' ? 'expiring' : 'search',
        from: formData.from || '',
      })
    );
  } catch (error) {
    console.error('Error updating customer:', error);
    res.writeHead(500, { 'Content-Type': 'text/html; charset=utf-8' });
    res.end(
      renderEditCustomerPage({
        message: 'เกิดข้อผิดพลาดในการบันทึกข้อมูล',
        status: 'error',
      })
    );
  }
}

/**
 * Handles bulk delete of customers
 */
export async function handleDeleteCustomers(
  req: IncomingMessage,
  res: ServerResponse
): Promise<void> {
  try {
    const parsed = await parseBody(req);
    const rowsRaw = parsed.rows;
    const rows = Array.isArray(rowsRaw) ? rowsRaw : rowsRaw ? [rowsRaw] : [];
    const rowNumbers = rows
      .map((r) => Number.parseInt(String(r), 10))
      .filter((n) => Number.isFinite(n) && n >= 2);

    if (rowNumbers.length === 0) {
      res.writeHead(302, { Location: '/customers/search' });
      res.end();
      return;
    }

    await deleteCustomersInSheet(rowNumbers);
    res.writeHead(302, { Location: '/customers/search' });
    res.end();
  } catch (error) {
    console.error('Error deleting customers:', error);
    res.writeHead(500, { 'Content-Type': 'text/html; charset=utf-8' });
    res.end('เกิดข้อผิดพลาดในการลบข้อมูล');
  }
}
