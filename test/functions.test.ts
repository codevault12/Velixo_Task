/* eslint-disable @typescript-eslint/no-explicit-any */
import { factorialSeries } from "../src/shared/factorial";

declare const require: (moduleId: string) => any;

// Stand-in for CustomFunctions.Error so validation failures are assertable.
class MockCustomFunctionError extends Error {
  constructor(
    public code: string,
    message?: string
  ) {
    super(message);
    this.name = "CustomFunctionError";
  }
}

type FactorialRow = (n: number) => Promise<Array<Array<number | string>>>;

describe("factorialRow (custom function)", () => {
  let store: Record<string, string>;

  beforeEach(() => {
    jest.resetModules();
    store = {};
    delete (window as { __velixo?: unknown }).__velixo;

    (globalThis as any).CustomFunctions = {
      associate: jest.fn(),
      Error: MockCustomFunctionError,
      ErrorCode: { invalidValue: "invalidValue" },
    };
    (globalThis as any).OfficeRuntime = {
      storage: {
        getItem: (key: string) => Promise.resolve(store[key] ?? null),
        setItem: (key: string, value: string) => {
          store[key] = value;
          return Promise.resolve();
        },
      },
    };
  });

  // Loads a fresh module graph with the web worker replaced by an in-process
  // computation, so the function's own logic is exercised without a real Worker.
  function loadFunction(): FactorialRow {
    jest.doMock("../src/functions/worker-client", () => ({
      computeFactorials: (n: number) =>
        Promise.resolve(factorialSeries(n, [1n]).map((value) => value.toString())),
    }));
    return require("../src/functions/functions").factorialRow;
  }

  it("returns a single row of 0!..N! by default", async () => {
    const factorialRow = loadFunction();
    await expect(factorialRow(5)).resolves.toEqual([[1, 1, 2, 6, 24, 120]]);
  });

  it("returns 0! as a single cell for N = 0", async () => {
    const factorialRow = loadFunction();
    await expect(factorialRow(0)).resolves.toEqual([[1]]);
  });

  it("returns a column when the persisted orientation is column", async () => {
    store["velixo:orientation"] = "column";
    const factorialRow = loadFunction();
    await expect(factorialRow(3)).resolves.toEqual([[1], [1], [2], [6]]);
  });

  it("returns large factorials as lossless strings", async () => {
    const factorialRow = loadFunction();
    const [row] = await factorialRow(20);
    expect(row[18]).toBe(6402373705728000); // 18! still a number
    expect(row[19]).toBe("121645100408832000"); // 19! becomes a string
    expect(row[20]).toBe("2432902008176640000"); // 20!
  });

  it("rejects out-of-range or non-integer N with a custom-function error", async () => {
    const factorialRow = loadFunction();
    await expect(factorialRow(-1)).rejects.toBeInstanceOf(MockCustomFunctionError);
    await expect(factorialRow(501)).rejects.toBeInstanceOf(MockCustomFunctionError);
    await expect(factorialRow(2.5)).rejects.toBeInstanceOf(MockCustomFunctionError);
  });
});
