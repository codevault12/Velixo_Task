/// <reference lib="webworker" />
import { factorialSeries } from "../shared/factorial";
import type { FactorialRequest, FactorialResponse } from "../shared/worker-protocol";

// Persistent cache: cache[i] === i!. Seeded with 0! === 1 and only ever grown,
// so any individual factorial is computed exactly once per worker lifetime.
const cache: bigint[] = [1n];

const ctx = self as unknown as DedicatedWorkerGlobalScope;

ctx.onmessage = (event: MessageEvent<FactorialRequest>) => {
  const { jobId, n } = event.data;
  const series = factorialSeries(n, cache);
  const response: FactorialResponse = {
    jobId,
    values: series.map((value) => value.toString()),
  };
  ctx.postMessage(response);
};
