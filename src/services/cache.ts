import { Customer } from '../types/index.js';

/**
 * Simple LRU Cache implementation for customer data
 * This significantly improves performance by reducing API calls
 */
class CustomerCache {
  private cache: Customer[] | null = null;
  private lastFetch: number = 0;
  private readonly TTL: number = 30000; // 30 seconds cache TTL

  /**
   * Get cached customers if still valid
   */
  get(): Customer[] | null {
    if (this.cache && Date.now() - this.lastFetch < this.TTL) {
      return this.cache;
    }
    return null;
  }

  /**
   * Set customers in cache
   */
  set(customers: Customer[]): void {
    this.cache = customers;
    this.lastFetch = Date.now();
  }

  /**
   * Invalidate the cache (call this after create/update/delete operations)
   */
  invalidate(): void {
    this.cache = null;
    this.lastFetch = 0;
  }

  /**
   * Check if cache is still valid
   */
  isValid(): boolean {
    return this.cache !== null && Date.now() - this.lastFetch < this.TTL;
  }
}

export const customerCache = new CustomerCache();
