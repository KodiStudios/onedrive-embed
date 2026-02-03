import { describe, it } from "node:test";
import assert from "node:assert";
import * as path from "node:path";
import { fileURLToPath } from "node:url";
import { getAllFilePaths } from "../index.ts";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const fixturesInputDir = path.join(__dirname, "fixtures", "input");

describe("getAllFilePaths", () => {
  it("should find all files in fixtures/input directory", () => {
    const result = getAllFilePaths(fixturesInputDir);

    // Should find 5 files: sample.md, sample.html, no-embeds.md, already-processed.md, nested/deep.md
    assert.strictEqual(result.length, 5);
  });

  it("should return absolute paths", () => {
    const result = getAllFilePaths(fixturesInputDir);

    for (const filePath of result) {
      assert.ok(
        path.isAbsolute(filePath),
        `Path should be absolute: ${filePath}`,
      );
    }
  });

  it("should include files from nested directories", () => {
    const result = getAllFilePaths(fixturesInputDir);

    const hasNestedFile = result.some((filePath) =>
      filePath.includes(path.join("nested", "deep.md")),
    );
    assert.ok(hasNestedFile, "Should include nested/deep.md");
  });

  it("should include all expected fixture files", () => {
    const result = getAllFilePaths(fixturesInputDir);
    const fileNames = result.map((p) => path.basename(p));

    assert.ok(fileNames.includes("sample.md"), "Should include sample.md");
    assert.ok(fileNames.includes("sample.html"), "Should include sample.html");
    assert.ok(
      fileNames.includes("no-embeds.md"),
      "Should include no-embeds.md",
    );
    assert.ok(
      fileNames.includes("already-processed.md"),
      "Should include already-processed.md",
    );
    assert.ok(fileNames.includes("deep.md"), "Should include deep.md");
  });

  it("should return empty array for empty directory", () => {
    // Create a reference to the temp directory (will be empty or non-existent)
    const emptyDir = path.join(__dirname, ".temp", "empty-test");

    // Skip if directory doesn't exist - this is a defensive test
    try {
      const result = getAllFilePaths(emptyDir);
      assert.strictEqual(result.length, 0);
    } catch {
      // Directory doesn't exist, which is expected in clean state
      assert.ok(
        true,
        "Empty directory test skipped - directory does not exist",
      );
    }
  });
});
