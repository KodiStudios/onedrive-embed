import { describe, it } from "node:test";
import assert from "node:assert";
import { getOneDriveFilePath } from "../index.ts";

describe("getOneDriveFilePath", () => {
  it("should construct correct path from Graph API response", async (t) => {
    // Track which API paths were called
    const apiCalls: string[] = [];

    const mockGraphClient = {
      api: (apiPath: string) => {
        apiCalls.push(apiPath);
        return {
          get: async () => {
            if (apiPath.includes("/shares/")) {
              // First call: /shares/{sharedItemId}/driveItem
              return { id: "test-item-id-123" };
            } else {
              // Second call: /me/drive/items/{itemId}
              return {
                name: "photo.png",
                parentReference: {
                  path: "/drive/root:/Pictures/vacation",
                },
              };
            }
          },
        };
      },
    };

    const result = await getOneDriveFilePath(
      "s!TestSharedId",
      mockGraphClient as any
    );

    assert.strictEqual(result, "/Pictures/vacation/photo.png");
    assert.strictEqual(apiCalls.length, 2);
    assert.ok(apiCalls[0].includes("/shares/s!TestSharedId/driveItem"));
    assert.ok(apiCalls[1].includes("/me/drive/items/test-item-id-123"));
  });

  it("should handle nested folder paths", async () => {
    const mockGraphClient = {
      api: (apiPath: string) => ({
        get: async () => {
          if (apiPath.includes("/shares/")) {
            return { id: "nested-item-id" };
          } else {
            return {
              name: "deep-image.jpg",
              parentReference: {
                path: "/drive/root:/Photos/2024/summer/beach",
              },
            };
          }
        },
      }),
    };

    const result = await getOneDriveFilePath(
      "s!NestedSharedId",
      mockGraphClient as any
    );

    assert.strictEqual(result, "/Photos/2024/summer/beach/deep-image.jpg");
  });

  it("should handle root-level files", async () => {
    const mockGraphClient = {
      api: (apiPath: string) => ({
        get: async () => {
          if (apiPath.includes("/shares/")) {
            return { id: "root-item-id" };
          } else {
            return {
              name: "root-file.png",
              parentReference: {
                path: "/drive/root:",
              },
            };
          }
        },
      }),
    };

    const result = await getOneDriveFilePath(
      "s!RootSharedId",
      mockGraphClient as any
    );

    // path.posix.join with empty string doesn't prepend /
    assert.strictEqual(result, "root-file.png");
  });

  it("should handle files with special characters in name", async () => {
    const mockGraphClient = {
      api: (apiPath: string) => ({
        get: async () => {
          if (apiPath.includes("/shares/")) {
            return { id: "special-item-id" };
          } else {
            return {
              name: "photo (1).png",
              parentReference: {
                path: "/drive/root:/My Pictures",
              },
            };
          }
        },
      }),
    };

    const result = await getOneDriveFilePath(
      "s!SpecialSharedId",
      mockGraphClient as any
    );

    assert.strictEqual(result, "/My Pictures/photo (1).png");
  });
});
