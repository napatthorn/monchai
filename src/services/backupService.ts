import { writeFile, mkdir } from 'fs/promises';
import { readdirSync } from 'fs';
import { dirname } from 'path';
import { fileURLToPath } from 'url';
import { Customer } from '../types/index.js';
import { fetchCustomers } from './customerService.js';

const __dirname = dirname(fileURLToPath(import.meta.url));
const BACKUP_DIR = `${__dirname}/../../data`;

/**
 * Saves customers data to a timestamped JSON file
 */
export async function saveCustomersToFile(customers: Customer[]): Promise<string> {
  await mkdir(BACKUP_DIR, { recursive: true });
  const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
  const filename = `customers_${timestamp}.json`;
  const filepath = `${BACKUP_DIR}/${filename}`;
  
  await writeFile(filepath, JSON.stringify(customers, null, 2), 'utf-8');
  return filepath;
}

/**
 * Runs backup on server startup (once per day)
 */
export async function runBackupOnStartup(): Promise<void> {
  try {
    await mkdir(BACKUP_DIR, { recursive: true });
    
    // Check if backup already exists for today
    const today = new Date().toISOString().split('T')[0];
    const backupPattern = new RegExp(`customers_${today}`);
    
    const backupExists = readdirSync(BACKUP_DIR)
      .some(file => backupPattern.test(file));

    if (backupExists) {
      console.log('Backup already exists for today, skipping...');
      return;
    }

    console.log('Running startup backup...');
    const customers = await fetchCustomers();
    const backupPath = await saveCustomersToFile(customers);
    console.log(`Backup created successfully: ${backupPath}`);
  } catch (error) {
    console.error('Error during startup backup:', error);
  }
}
