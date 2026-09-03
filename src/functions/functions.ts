import { toCellValue } from "../shared/factorial";
import { getOrientation, loadOrientation, toSpillRange, type Orientation } from "../shared/orientation";
import { computeFactorials } from "./worker-client";

const MAX_N = 500;

// Hydrate the persisted orientation once, lazily, on the first call. This lets
// the persisted setting apply even before the task pane has been opened.
let orientationReady: Promise<Orientation> | undefined;
function ensureOrientationLoaded(): Promise<Orientation> {
  return (orientationReady ??= loadOrientation());
}

/**
 * Returns the factorials from 0! to N! as a spill range.
 *
 * The orientation (row or column) follows the setting on the add-in's task
 * pane. Values that exceed JavaScript's safe-integer range are returned as
 * exact string representations, so results stay lossless up to N = 500.
 * @customfunction FACTORIALROW
 * @param n Upper bound N (an integer between 0 and 500).
 * @returns A spill range containing [0!, 1!, 2!, ..., N!].
 */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
export async function factorialRow(n: number): Promise<any[][]> {
  if (!Number.isInteger(n) || n < 0 || n > MAX_N) {
    throw new CustomFunctions.Error(
      CustomFunctions.ErrorCode.invalidValue,
      `N must be an integer between 0 and ${MAX_N}.`
    );
  }

  await ensureOrientationLoaded();
  const values = (await computeFactorials(n)).map((value) => toCellValue(BigInt(value)));
  return toSpillRange(values, getOrientation());
}

CustomFunctions.associate("FACTORIALROW", factorialRow);
