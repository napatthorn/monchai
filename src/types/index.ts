export interface Customer {
  rowNumber?: number;
  timestamp: string;
  customerName: string;
  licensePlate: string;
  policyNumber: string | null;
  actIssuedDate: string | null;
  actExpiryDate: string | null;
  taxRenewalDate: string | null;
  taxExpiryDate: string | null;
  voluntaryIssuedDate: string | null;
  voluntaryExpiryDate: string | null;
  registrationDate: string | null;
  email: string | null;
  phone: string;
  status: string | null;
  notes: string | null;
}

export interface FormData {
  timestamp?: string;
  customerName: string;
  licensePlate: string;
  policyNumber: string;
  actIssuedDate: string;
  actExpiryDate: string;
  taxRenewalDate: string;
  taxExpiryDate: string;
  voluntaryIssuedDate: string;
  voluntaryExpiryDate: string;
  registrationDate: string;
  email: string;
  phone: string;
  status: string;
  notes: string;
  rowNumber?: string;
  from?: string;
}

export interface ValidationErrors {
  [key: string]: string;
}

export interface RenderOptions {
  pageTitle: string;
  active?: string;
  content?: string;
  alertMarkup?: string;
}

export interface CustomerFormOptions {
  formData?: Partial<FormData>;
  errors?: ValidationErrors;
  message?: string;
  status?: 'success' | 'error';
  isEditMode?: boolean;
  showStatus?: boolean;
  activeNav?: string;
  from?: string;
}

export interface SearchPageOptions {
  query?: string;
  results?: Customer[];
  total?: number;
}

export interface ExpiringPageOptions {
  customers?: Customer[];
  days?: number;
}

export interface SheetResponse {
  records?: Customer[];
}

export interface Config {
  PORT: number;
  HOST: string;
  APP_TITLE: string;
  SHEET_WRITE_URL: string;
  SHEET_DATA_URL: string;
  SHEET_UPDATE_URL: string;
  DEFAULT_EXPIRY_WINDOW_DAYS: number;
}
