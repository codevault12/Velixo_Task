import type { FactorialRequest, FactorialResponse } from "../shared/worker-protocol";

/**
 * Thin client around a single, long-lived factorial web worker. The worker is
 * created lazily and reused, so its factorial cache persists across calls.
 * Requests are matched to responses by jobId so concurrent calls never cross.
 */

interface PendingJob {
  resolve: (values: string[]) => void;
  reject: (error: Error) => void;
}

let worker: Worker | undefined;
let nextJobId = 0;
const pending = new Map<number, PendingJob>();

function getWorker(): Worker {
  if (!worker) {
    worker = new Worker("factorial.worker.js");
    worker.onmessage = (event: MessageEvent<FactorialResponse>) => {
      const { jobId, values } = event.data;
      const job = pending.get(jobId);
      if (job) {
        pending.delete(jobId);
        job.resolve(values);
      }
    };
    // Fail every outstanding job if the worker errors, then clear the
    // singleton so the next call spawns a fresh worker instead of posting
    // to the dead one and hanging forever.
    worker.onerror = (event) => {
      const error = new Error(`Factorial worker failed: ${event.message}`);
      worker = undefined;
      for (const [jobId, job] of pending) {
        pending.delete(jobId);
        job.reject(error);
      }
    };
  }
  return worker;
}

/** Computes 0!..N! off the UI thread, returning lossless decimal strings. */
export function computeFactorials(n: number): Promise<string[]> {
  return new Promise((resolve, reject) => {
    const jobId = nextJobId++;
    pending.set(jobId, { resolve, reject });
    const request: FactorialRequest = { jobId, n };
    getWorker().postMessage(request);
  });
}
