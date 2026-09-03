import {
  getOrientation,
  loadOrientation,
  persistOrientation,
  setOrientation,
  toSpillRange,
} from "../src/shared/orientation";

describe("toSpillRange", () => {
  it("wraps values as a single row", () => {
    expect(toSpillRange([1, 2, 6], "row")).toEqual([[1, 2, 6]]);
  });

  it("wraps values as a single column", () => {
    expect(toSpillRange([1, 2, 6], "column")).toEqual([[1], [2], [6]]);
  });
});

describe("orientation state", () => {
  beforeEach(() => {
    delete (window as { __velixo?: unknown }).__velixo;
  });

  it("defaults to row when unset", () => {
    expect(getOrientation()).toBe("row");
  });

  it("reflects the in-memory value", () => {
    setOrientation("column");
    expect(getOrientation()).toBe("column");
  });
});

describe("orientation persistence", () => {
  let store: Record<string, string>;

  beforeEach(() => {
    store = {};
    delete (window as { __velixo?: unknown }).__velixo;
    (globalThis as unknown as { OfficeRuntime: unknown }).OfficeRuntime = {
      storage: {
        getItem: (key: string) => Promise.resolve(store[key] ?? null),
        setItem: (key: string, value: string) => {
          store[key] = value;
          return Promise.resolve();
        },
      },
    };
  });

  it("persists and reloads the chosen orientation", async () => {
    await persistOrientation("column");
    expect(store["velixo:orientation"]).toBe("column");

    delete (window as { __velixo?: unknown }).__velixo;
    await expect(loadOrientation()).resolves.toBe("column");
    expect(getOrientation()).toBe("column");
  });

  it("falls back to row for missing or unexpected stored values", async () => {
    await expect(loadOrientation()).resolves.toBe("row");
    store["velixo:orientation"] = "diagonal";
    await expect(loadOrientation()).resolves.toBe("row");
  });
});
