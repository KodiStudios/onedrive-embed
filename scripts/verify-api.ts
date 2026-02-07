import { Client } from "@microsoft/microsoft-graph-client";
import minimist from "minimist";
import { getOneDriveFilePath } from "../index.ts";

const argv = minimist(process.argv.slice(2));

if (!argv.token || !argv["shared-id"]) {
  console.log("Usage:");
  console.log(
    `  node scripts/test-api.ts --token {token_from_aka.ms/ge} --shared-id {shared_id}`,
  );
  console.log();
  console.log("Example:");
  console.log(
    `  node scripts/test-api.ts --token eyJ... --shared-id "s!AmslmcZf6z3Lg98-IHg6iib_9ykDOw"`,
  );
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
  const oneDriveFilePath = await getOneDriveFilePath(
    argv["shared-id"],
    graphClient,
  );
  console.log();
  console.log(`Resolved OneDrive path: ${oneDriveFilePath}`);
} catch (error) {
  console.error();
  console.error(`Failed to resolve shared ID: ${argv["shared-id"]}`);
  console.error(error);
  process.exit(1);
}
