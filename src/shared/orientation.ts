/**
 * Orientation setting shared between the task pane and the custom function.
 *
 * Under the shared runtime both bundles execute in the same `window`, but they
 * are separate webpack chunks and therefore do NOT share module-level state.
 * The single source of truth is `window.__velixo`, which both chunks read/write.
 * `OfficeRuntime.storage` backs it so the choice survives a page reload.
 */

export type Orientation = "row" | "column";

const STORAGE_KEY = "velixo:orientation";
const DEFAULT_ORIENTATION: Orientation = "row";

declare global {
  interface Window {
    __velixo?: { orientation?: Orientation };
  }
}

/** Reads the current orientation synchronously (used by the custom function). */
export function getOrientation(): Orientation {
  return window.__velixo?.orientation ?? DEFAULT_ORIENTATION;
}

/** Updates the in-memory orientation shared across the runtime. */
export function setOrientation(orientation: Orientation): void {
  (window.__velixo ??= {}).orientation = orientation;
}

/** Hydrates the in-memory orientation from persistent storage. */
export async function loadOrientation(): Promise<Orientation> {
  const stored = await OfficeRuntime.storage.getItem(STORAGE_KEY);
  const orientation: Orientation = stored === "column" ? "column" : DEFAULT_ORIENTATION;
  setOrientation(orientation);
  return orientation;
}

/** Persists the orientation and updates the in-memory value. */
export async function persistOrientation(orientation: Orientation): Promise<void> {
  setOrientation(orientation);
  await OfficeRuntime.storage.setItem(STORAGE_KEY, orientation);
}

/**
 * Shapes a flat list of values into a 2D spill range for Excel:
 * a single row (`[[a, b, c]]`) or a single column (`[[a], [b], [c]]`).
 */
export function toSpillRange<T>(values: T[], orientation: Orientation): T[][] {
  return orientation === "row" ? [values] : values.map((value) => [value]);
}
