import { factorial, factorialSeries, toCellValue } from "../src/shared/factorial";

describe("factorial", () => {
  it("computes small factorials exactly", () => {
    expect(factorial(0)).toBe(1n);
    expect(factorial(1)).toBe(1n);
    expect(factorial(5)).toBe(120n);
    expect(factorial(10)).toBe(3628800n);
  });

  it("stays exact well beyond Number precision (20!)", () => {
    expect(factorial(20)).toBe(2432902008176640000n);
  });

  it("has the correct number of digits for large N", () => {
    expect(factorial(100).toString()).toHaveLength(158);
    expect(factorial(500).toString()).toHaveLength(1135);
  });

  it("rejects negative and non-integer input", () => {
    expect(() => factorial(-1)).toThrow(RangeError);
    expect(() => factorial(1.5)).toThrow(RangeError);
  });
});

describe("factorialSeries", () => {
  it("returns the full series 0!..N!", () => {
    expect(factorialSeries(5, [1n])).toEqual([1n, 1n, 2n, 6n, 24n, 120n]);
  });

  it("reuses and extends a shared cache (each value computed once)", () => {
    const cache: bigint[] = [1n];
    factorialSeries(5, cache);
    expect(cache).toHaveLength(6);

    // A larger request extends the same cache without recomputing prior values.
    const five = cache[5];
    factorialSeries(8, cache);
    expect(cache).toHaveLength(9);
    expect(cache[5]).toBe(five);

    // A smaller request adds nothing.
    factorialSeries(3, cache);
    expect(cache).toHaveLength(9);
  });
});

describe("toCellValue", () => {
  it("returns a JS number while within safe-integer range (<= 18!)", () => {
    expect(toCellValue(factorial(18))).toBe(6402373705728000);
    expect(typeof toCellValue(factorial(18))).toBe("number");
  });

  it("returns a lossless string once the value exceeds safe integers (>= 19!)", () => {
    const value = toCellValue(factorial(19));
    expect(typeof value).toBe("string");
    expect(value).toBe("121645100408832000");
  });
});
