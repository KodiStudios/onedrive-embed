import { Client } from "@microsoft/microsoft-graph-client";
import fs from "fs";
import minimist from "minimist";
import path from "path";

function findFileSharedItemIds(
  filePath: string
): Array</*sharedItemId*/ string> {
  const sharedItemIds = new Array<string>();

  const fileContent: string = fs.readFileSync(filePath, `utf8`);

  // Sample OneDrive Shared Id element:
  // <img src="https://1drv.ms/i/s!AmslmcZf6z3Lg98-IHg6iib_9ykDOw?embed=1&width=981&height=740" width="981" height="740" />
  // Where Shared Id is:
  // s!AmslmcZf6z3Lg98-IHg6iib_9ykDOw
  //                                         : / /1drv.ms /i /SharedId
  for (const match of fileContent.matchAll(/\:\/\/1drv.ms\/i\/([^\?]+)/g)) {
    sharedItemIds.push(match[1]);
  }

  return sharedItemIds;
}

function findDirectorySharedItemIds(
  directoryPath: string
): Array</*sharedItemId*/ string> {
  const sharedItemIds = new Array<string>();

  const filesOrDirectoryNames = fs.readdirSync(directoryPath);

  for (const fileOrDirectoryName of filesOrDirectoryNames) {
    const fileOrDirectoryPath = path.join(directoryPath, fileOrDirectoryName);
    if (fs.statSync(fileOrDirectoryPath).isDirectory()) {
      // Directory
      sharedItemIds.push(...findDirectorySharedItemIds(fileOrDirectoryPath));
    } else {
      // File
      sharedItemIds.push(...findFileSharedItemIds(fileOrDirectoryPath));
    }
  }

  return sharedItemIds;
}

async function getOneDriveFilePath(
  sharedItemId: string,
  graphClient: Client
): Promise<string> {
  let sharedDriveItem: any = await graphClient
    .api(`/shares/${sharedItemId}/driveItem`)
    .get();

  let itemId: string = sharedDriveItem.id;
  console.log(`itemId: ${itemId}`);

  // Note that sharedDriveItem has different fields than driveItem
  // Only driveItem contains OneDrive directory path

  let driveItem: any = await graphClient.api(`/me/drive/items/${itemId}`).get();
  console.log(driveItem);

  // parentReferencePath Format:
  // /drive/root:/Pictures/metro-evolved-pictures
  //                                                        oot:/Pic
  const matchGroups = driveItem.parentReference.path.match(/.*\:(.*)/);
  const directoryPath: string = matchGroups![1];

  // driveItem.name is file name
  const filePath: string = path.join(directoryPath, driveItem.name);
  console.log(`filePath: ${filePath}`);

  return filePath;
}

async function main(): Promise<void> {
  const argv = minimist(process.argv.slice(2));
  if (!argv.directory || !argv.token) {
    console.log("Usage: ");
    console.log(
      `${process.argv[0]} --experimental-strip-types ${
        import.meta.filename
      } --directory {directorypath} --token {token_value_from_aka.ms/ge}`
    );
    return;
  }

  const sharedItemIds: Array<string> = findDirectorySharedItemIds(
    argv.directory
  );

  const graphClient: Client = Client.init({
    defaultVersion: "v1.0",
    debugLogging: true, // Logs each Microsoft Graph query
    authProvider: (done) => {
      const errorMessage = "error throw by the authentication handler";
      done(errorMessage, argv.token);
    },
  });

  for (const sharedItemId of sharedItemIds) {
    console.log(`sharedItemId: ${sharedItemId}`);
    const oneDriveFilePath = await getOneDriveFilePath(
      sharedItemId,
      graphClient
    );
    console.log(`sharedItemId: ${sharedItemId}`);
    console.log(`oneDriveFilePath: ${oneDriveFilePath}`);
  }
}

await main();
