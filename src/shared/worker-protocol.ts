/** Message contract between the custom function and the factorial web worker. */

export interface FactorialRequest {
  jobId: number;
  /** Upper bound N; the worker returns 0! through N!. */
  n: number;
}

export interface FactorialResponse {
  jobId: number;
  /** Factorials 0!..N! as lossless decimal strings (BigInt is not always structured-cloneable). */
  values: string[];
}
