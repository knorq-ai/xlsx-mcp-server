import { describe, it, expect } from "vitest";
import { withFileLock } from "../engine/file-lock.js";

describe("withFileLock", () => {
  it("serializes writes to the same file", async () => {
    const order: number[] = [];

    const p1 = withFileLock("/tmp/test-lock.xlsx", async () => {
      await new Promise((r) => setTimeout(r, 50));
      order.push(1);
    });

    const p2 = withFileLock("/tmp/test-lock.xlsx", async () => {
      order.push(2);
    });

    await Promise.all([p1, p2]);
    expect(order).toEqual([1, 2]);
  });

  it("allows parallel writes to different files", async () => {
    const order: string[] = [];

    const p1 = withFileLock("/tmp/test-lock-a.xlsx", async () => {
      await new Promise((r) => setTimeout(r, 50));
      order.push("a");
    });

    const p2 = withFileLock("/tmp/test-lock-b.xlsx", async () => {
      order.push("b");
    });

    await Promise.all([p1, p2]);
    // b should finish before a
    expect(order).toEqual(["b", "a"]);
  });

  it("releases lock on error", async () => {
    try {
      await withFileLock("/tmp/test-lock-err.xlsx", async () => {
        throw new Error("test error");
      });
    } catch {
      // expected
    }

    // Should not deadlock
    let ran = false;
    await withFileLock("/tmp/test-lock-err.xlsx", async () => {
      ran = true;
    });
    expect(ran).toBe(true);
  });

  it("returns the value from the callback", async () => {
    const result = await withFileLock("/tmp/test-lock-ret.xlsx", async () => {
      return 42;
    });
    expect(result).toBe(42);
  });
});
