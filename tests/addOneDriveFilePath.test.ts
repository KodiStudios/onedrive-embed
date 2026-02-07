import { describe, it, before, after } from "node:test";
import assert from "node:assert";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { addOneDriveFilePath } from "../index.ts";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const fixturesDir = path.join(__dirname, "fixtures");
const inputDir = path.join(fixturesDir, "input");
const expectedDir = path.join(fixturesDir, "expected");
const tempDir = path.join(__dirname, ".temp");

/**
 * Recursively copy a directory
 */
function copyDirSync(src: string, dest: string): void {
  fs.mkdirSync(dest, { recursive: true });
  const entries = fs.readdirSync(src, { withFileTypes: true });

  for (const entry of entries) {
    const srcPath = path.join(src, entry.name);
    const destPath = path.join(dest, entry.name);

    if (entry.isDirectory()) {
      copyDirSync(srcPath, destPath);
    } else {
      fs.copyFileSync(srcPath, destPath);
    }
  }
}

/**
 * Recursively remove a directory
 */
function removeDirSync(dir: string): void {
  if (fs.existsSync(dir)) {
    fs.rmSync(dir, { recursive: true, force: true });
  }
}

/**
 * Create the shared ID to OneDrive path mapping for test fixtures
 */
function createTestSharedIdMap(): Map<string, string> {
  return new Map([
    // sample.md shared IDs
    ["s!AmslmcZf6z3Lg98-IHg6iib_9ykDOw", "/Pictures/photo1.png"],
    ["s!BnrtpcYg7a4Mh12-KJi7jkc_8zlEPx", "/Documents/image2.jpg"],
    ["s!CprumcZf6z3Lg99-JHg6iib_9ykDPw", "/Photos/vacation/sunset.png"],
    // sample.html shared IDs
    ["s!DqsuncZf6z3Lg00-LHg6iib_9ykDQw", "/Gallery/art1.png"],
    ["s!ErtvodZf6z3Lg01-MHg6iib_9ykDRw", "/Gallery/art2.png"],
    // nested/deep.md shared ID
    ["s!FsuwpeZf6z3Lg02-NHg6iib_9ykDSw", "/Archive/2024/deep-image.png"],
    // already-processed.md shared ID (should not be modified again)
    ["s!GtxwqfZf6z3Lg03-OHg6iib_9ykDTw", "/Pictures/already-tagged.png"],
  ]);
}

describe("addOneDriveFilePath", () => {
  let testRunDir: string;

  before(() => {
    // Clean up any previous test runs
    removeDirSync(tempDir);
    // Create temp directory
    fs.mkdirSync(tempDir, { recursive: true });
  });

  after(() => {
    // Clean up after all tests (comment out to inspect results)
    // removeDirSync(tempDir);
  });

  it("should add OneDrive path to markdown file with multiple embeds", () => {
    testRunDir = path.join(tempDir, "test-sample-md");
    fs.mkdirSync(testRunDir, { recursive: true });

    // Copy input file to temp
    const inputFile = path.join(inputDir, "sample.md");
    const tempFile = path.join(testRunDir, "sample.md");
    fs.copyFileSync(inputFile, tempFile);

    // Run the function
    const sharedIdMap = createTestSharedIdMap();
    addOneDriveFilePath(tempFile, sharedIdMap);

    // Compare with expected output
    const actualContent = fs.readFileSync(tempFile, "utf8");
    const expectedContent = fs.readFileSync(
      path.join(expectedDir, "sample.md"),
      "utf8",
    );

    assert.strictEqual(actualContent, expectedContent);
  });

  it("should add OneDrive path to HTML file", () => {
    testRunDir = path.join(tempDir, "test-sample-html");
    fs.mkdirSync(testRunDir, { recursive: true });

    const inputFile = path.join(inputDir, "sample.html");
    const tempFile = path.join(testRunDir, "sample.html");
    fs.copyFileSync(inputFile, tempFile);

    const sharedIdMap = createTestSharedIdMap();
    addOneDriveFilePath(tempFile, sharedIdMap);

    const actualContent = fs.readFileSync(tempFile, "utf8");
    const expectedContent = fs.readFileSync(
      path.join(expectedDir, "sample.html"),
      "utf8",
    );

    assert.strictEqual(actualContent, expectedContent);
  });

  it("should not modify file without OneDrive embeds", () => {
    testRunDir = path.join(tempDir, "test-no-embeds");
    fs.mkdirSync(testRunDir, { recursive: true });

    const inputFile = path.join(inputDir, "no-embeds.md");
    const tempFile = path.join(testRunDir, "no-embeds.md");
    fs.copyFileSync(inputFile, tempFile);

    // Get original content and modification time
    const originalContent = fs.readFileSync(tempFile, "utf8");

    const sharedIdMap = createTestSharedIdMap();
    addOneDriveFilePath(tempFile, sharedIdMap);

    // Content should be unchanged
    const actualContent = fs.readFileSync(tempFile, "utf8");
    assert.strictEqual(actualContent, originalContent);
  });

  it("should not modify already-processed URLs (URLs with hash fragments)", () => {
    testRunDir = path.join(tempDir, "test-already-processed");
    fs.mkdirSync(testRunDir, { recursive: true });

    const inputFile = path.join(inputDir, "already-processed.md");
    const tempFile = path.join(testRunDir, "already-processed.md");
    fs.copyFileSync(inputFile, tempFile);

    const originalContent = fs.readFileSync(tempFile, "utf8");

    const sharedIdMap = createTestSharedIdMap();
    addOneDriveFilePath(tempFile, sharedIdMap);

    // Content should be unchanged (already has hash fragment)
    const actualContent = fs.readFileSync(tempFile, "utf8");
    assert.strictEqual(actualContent, originalContent);
  });

  it("should handle nested directory files", () => {
    testRunDir = path.join(tempDir, "test-nested");
    fs.mkdirSync(path.join(testRunDir, "nested"), { recursive: true });

    const inputFile = path.join(inputDir, "nested", "deep.md");
    const tempFile = path.join(testRunDir, "nested", "deep.md");
    fs.copyFileSync(inputFile, tempFile);

    const sharedIdMap = createTestSharedIdMap();
    addOneDriveFilePath(tempFile, sharedIdMap);

    const actualContent = fs.readFileSync(tempFile, "utf8");
    const expectedContent = fs.readFileSync(
      path.join(expectedDir, "nested", "deep.md"),
      "utf8",
    );

    assert.strictEqual(actualContent, expectedContent);
  });

  it("should not modify file when shared ID is not in map", () => {
    testRunDir = path.join(tempDir, "test-missing-id");
    fs.mkdirSync(testRunDir, { recursive: true });

    const inputFile = path.join(inputDir, "sample.md");
    const tempFile = path.join(testRunDir, "sample.md");
    fs.copyFileSync(inputFile, tempFile);

    const originalContent = fs.readFileSync(tempFile, "utf8");

    // Empty map - no shared IDs known
    const emptyMap = new Map<string, string>();
    addOneDriveFilePath(tempFile, emptyMap);

    // Content should be unchanged
    const actualContent = fs.readFileSync(tempFile, "utf8");
    assert.strictEqual(actualContent, originalContent);
  });
});
