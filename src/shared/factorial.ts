/**
 * Pure factorial helpers. No Office/DOM/worker dependencies so they can be
 * unit-tested directly and reused by the web worker.
 */

/** Largest N whose factorial still fits in Number.MAX_SAFE_INTEGER (18! is safe, 19! is not). */
const MAX_SAFE_FACTORIAL_INPUT = 18;

/**
 * Extends `cache` in place so that `cache[i] === i!` for every `0 <= i <= n`,
 * then returns the slice `[0!, 1!, ..., n!]`.
 *
 * The caller owns the cache, seeded with `[1n]` (0! === 1). Because the cache
 * only ever grows, every individual factorial is multiplied exactly once for
 * the lifetime of that cache — a later smaller request costs zero work.
 */
export function factorialSeries(n: number, cache: bigint[]): bigint[] {
  if (!Number.isInteger(n) || n < 0) {
    throw new RangeError(`N must be a non-negative integer, received: ${n}`);
  }
  for (let i = cache.length; i <= n; i++) {
    cache[i] = cache[i - 1] * BigInt(i);
  }
  return cache.slice(0, n + 1);
}

/** Computes a single factorial with a throwaway cache. Convenience for tests. */
export function factorial(n: number): bigint {
  return factorialSeries(n, [1n])[n];
}

/**
 * Renders a factorial for an Excel cell: a JS number while it stays within
 * safe-integer range (so Excel treats it as a numeric value), otherwise a
 * lossless string representation.
 */
export function toCellValue(value: bigint): number | string {
  return value <= BigInt(Number.MAX_SAFE_INTEGER) ? Number(value) : value.toString();
}

export { MAX_SAFE_FACTORIAL_INPUT };
