import { RateLimitError } from '../errors';

export class RateLimiter {
  private requests: Map<string, number[]> = new Map();
  private limit: number;
  private window: number;

  constructor(limit: number = 60, windowMs: number = 60000) {
    this.limit = limit;
    this.window = windowMs;
  }

  async checkLimit(key: string): Promise<void> {
    const now = Date.now();
    const timestamps = this.requests.get(key) || [];
    
    // Remove old timestamps
    const validTimestamps = timestamps.filter(time => now - time < this.window);
    
    if (validTimestamps.length >= this.limit) {
      throw new RateLimitError('Rate limit exceeded. Please try again later.');
    }
    
    validTimestamps.push(now);
    this.requests.set(key, validTimestamps);
  }

  clear(): void {
    this.requests.clear();
  }
}