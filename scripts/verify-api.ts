import { Client } from "@microsoft/microsoft-graph-client";
import minimist from "minimist";
import { encodeSharingId, getOneDriveFilePath } from "../index.ts";

const argv = minimist(process.argv.slice(2));

// Default shared ID for testing (extracted from URL after /i/):
// Here's Url that is generated with "Embed"
// https://1drv.ms/i/c/fa8a0c5ee30de8e9/IQT0GKiU7k1iR7Adl607KrOfAarhD12HBp17ijxsIWQD2ss
const defaultSharedId =
  "c/fa8a0c5ee30de8e9/IQT0GKiU7k1iR7Adl607KrOfAarhD12HBp17ijxsIWQD2ss";

// Here's Url that is generated with "Share"
// https://1drv.ms/i/c/fa8a0c5ee30de8e9/IQD0GKiU7k1iR7Adl607KrOfAetxTji6uqLZIPs5Ms171vo?e=55Jrso
// Suffix "?e=55Jrso" is not relevant, it's not part of shared ID and should be stripped out
// by regex in extractSharedItemIdsFromFileContent()

const sharedId: string = argv["shared-id"] ?? defaultSharedId;

if (!argv.token) {
  console.log("Usage:");
  console.log(
    `  node scripts/verify-api.ts --token {token_from_aka.ms/ge} [--shared-id {shared_id}]`,
  );
  console.log();
  console.log("Example:");
  console.log(`  node scripts/verify-api.ts --token eyJ...`);
  process.exit(1);
}

const graphClient: Client = Client.init({
  defaultVersion: "v1.0",
  debugLogging: true,
  authProvider: (done) => {
    done("error throw by the authentication handler", argv.token);
  },
});

try {
  const oneDriveFilePath = await getOneDriveFilePath(sharedId, graphClient);
  console.log();
  console.log(`Resolved OneDrive path: ${oneDriveFilePath}`);
} catch (error) {
  console.error();
  console.error(`Failed to resolve shared ID: ${sharedId}`);
  console.error(error);
  process.exit(1);
}
