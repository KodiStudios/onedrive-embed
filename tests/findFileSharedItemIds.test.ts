import { describe, it } from "node:test";
import assert from "node:assert";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { findFileSharedItemIds } from "../index.ts";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const fixturesDir = path.join(__dirname, "fixtures", "input");

describe("findFileSharedItemIds", () => {
  it("should find multiple shared IDs from markdown file", () => {
    const filePath = path.join(fixturesDir, "sample.md");
    const result = findFileSharedItemIds(filePath);

    assert.strictEqual(result.size, 3);
    assert.ok(result.has("s!AmslmcZf6z3Lg98-IHg6iib_9ykDOw"));
    assert.ok(result.has("s!BnrtpcYg7a4Mh12-KJi7jkc_8zlEPx"));
    assert.ok(result.has("s!CprumcZf6z3Lg99-JHg6iib_9ykDPw"));
  });

  it("should find shared IDs from HTML file", () => {
    const filePath = path.join(fixturesDir, "sample.html");
    const result = findFileSharedItemIds(filePath);

    assert.strictEqual(result.size, 2);
    assert.ok(result.has("s!DqsuncZf6z3Lg00-LHg6iib_9ykDQw"));
    assert.ok(result.has("s!ErtvodZf6z3Lg01-MHg6iib_9ykDRw"));
  });

  it("should return empty set when no OneDrive embeds found", () => {
    const filePath = path.join(fixturesDir, "no-embeds.md");
    const result = findFileSharedItemIds(filePath);

    assert.strictEqual(result.size, 0);
  });

  it("should find shared IDs from nested files", () => {
    const filePath = path.join(fixturesDir, "nested", "deep.md");
    const result = findFileSharedItemIds(filePath);

    assert.strictEqual(result.size, 1);
    assert.ok(result.has("s!FsuwpeZf6z3Lg02-NHg6iib_9ykDSw"));
  });

  it("should find shared ID from already-processed file (ID is still extractable)", () => {
    const filePath = path.join(fixturesDir, "already-processed.md");
    const result = findFileSharedItemIds(filePath);

    // The shared ID is still present in the URL, even with the hash fragment
    assert.strictEqual(result.size, 1);
    assert.ok(result.has("s!GtxwqfZf6z3Lg03-OHg6iib_9ykDTw"));
  });
});
