import 'dotenv/config';
import { Config } from './types/index.js';

export const config: Config = {
  PORT: parseInt(process.env.PORT || '3000', 10),
  HOST: '127.0.0.1',
  APP_TITLE: 'Monchai Insurance',
  SHEET_WRITE_URL: process.env.SHEET_WEBHOOK_URL || '',
  SHEET_DATA_URL: process.env.SHEET_DATA_URL || '',
  SHEET_UPDATE_URL: process.env.SHEET_UPDATE_URL || process.env.SHEET_WEBHOOK_URL || '',
  DEFAULT_EXPIRY_WINDOW_DAYS: 30,
};
