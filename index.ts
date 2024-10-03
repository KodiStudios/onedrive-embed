import { Client } from "@microsoft/microsoft-graph-client";
import fs, { Dirent } from "node:fs";
import minimist from "minimist";
import path from "node:path";

function addOneDriveFilePath(
  filePath: string,
  sharedIdOdFileHash: Map</*sharedId*/ string, /*oneDrivePath*/ string>
) {
  const fileContent: string = fs.readFileSync(filePath, `utf8`);
  // <img src="https://1drv.ms/i/s!AmslmcZf6z3Lg98-IHg6iib_9ykDOw?embed=1&width=981&height=740" width="981" height="740" />

  //         ttps: / /1drv.ms /i /s!Ase"
  let reg = /".*\:\/\/1drv.ms\/i\/[^"]+"/g;

  let contentChanged = false;
  const updatedFileContent = fileContent.replace(
    reg,
    /*replacer*/ (quotedUrl: string, ...args: any[]): string => {
      console.log(`Matched quotedUrl: ${quotedUrl}`);
      let resultString = quotedUrl;
      if (quotedUrl.match(/\#/)) {
        // Already has one, Noop
        console.log(`Already Has #`);
      } else {
        let matches: RegExpMatchArray | null = quotedUrl.match(
          /\:\/\/1drv.ms\/i\/([^\?]+)/
        );
        if (matches) {
          // Get Shared Id
          const sharedId3: string = matches[1];
          console.log(`Matched SharedId: ${sharedId3}`);
          const oneDriveFilePath: string | undefined =
            sharedIdOdFileHash.get(sharedId3);
          if (oneDriveFilePath) {
            resultString = quotedUrl.replace(/"$/, `#${oneDriveFilePath}"`);
            console.log(`Updated: ${resultString}`);
            contentChanged = true;
          }
        } else {
          console.log(`Can't find SharedId`);
        }
      }

      return resultString;
    }
  );

  if (contentChanged) {
    console.log(`Writing File: ${filePath}`);
    fs.writeFileSync(filePath, updatedFileContent, {
      encoding: "utf8",
      flag: "w",
    });
  }
}

function findFileSharedItemIds(filePath: string): Set</*sharedItemId*/ string> {
  const sharedItemIds = new Set<string>();

  const fileContent: string = fs.readFileSync(filePath, `utf8`);

  // Sample OneDrive Shared Id element:
  // <img src="https://1drv.ms/i/s!AmslmcZf6z3Lg98-IHg6iib_9ykDOw?embed=1&width=981&height=740" width="981" height="740" />
  // Where Shared Id is:
  // s!AmslmcZf6z3Lg98-IHg6iib_9ykDOw
  //                                         : / /1drv.ms /i /SharedId
  for (const match of fileContent.matchAll(/\:\/\/1drv.ms\/i\/([^\?]+)/g)) {
    sharedItemIds.add(match[1]);
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
  const oneDriveFilePath: string = path.posix.join(
    directoryPath,
    driveItem.name
  );
  console.log(`filePath: ${oneDriveFilePath}`);

  return oneDriveFilePath;
}

function getAllFilePaths(directoryPath: string): Array<string> {
  let filePaths = new Array<string>();
  const fileOrDirectories: Dirent[] = fs.readdirSync(directoryPath, {
    withFileTypes: true,
  });

  for (const fileOrDirectory of fileOrDirectories) {
    const fileOrDirectoryPath = path.join(directoryPath, fileOrDirectory.name);
    if (fileOrDirectory.isDirectory()) {
      // Directory
      filePaths.push(...getAllFilePaths(fileOrDirectoryPath));
    } else {
      // File
      filePaths.push(fileOrDirectoryPath);
    }
  }

  return filePaths;
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

  const filePaths: Array<string> = getAllFilePaths(argv.directory);

  const sharedItemIds = new Set<string>();
  for (const filePath of filePaths) {
    findFileSharedItemIds(filePath).forEach((value) =>
      sharedItemIds.add(value)
    );
  }

  const graphClient: Client = Client.init({
    defaultVersion: "v1.0",
    debugLogging: true, // Logs each Microsoft Graph query
    authProvider: (done) => {
      const errorMessage = "error throw by the authentication handler";
      done(errorMessage, argv.token);
    },
  });

  let oneDriveFilePathMap = new Map<
    /*sharedItemId*/ string,
    /*filePath*/ string
  >();

  for (const sharedItemId of sharedItemIds) {
    console.log(`sharedItemId: ${sharedItemId}`);
    const oneDriveFilePath = await getOneDriveFilePath(
      sharedItemId,
      graphClient
    );
    console.log(`sharedItemId: ${sharedItemId}`);
    console.log(`oneDriveFilePath: ${oneDriveFilePath}`);

    oneDriveFilePathMap.set(sharedItemId, oneDriveFilePath);
  }

  for (const filePath of filePaths) {
    addOneDriveFilePath(filePath, oneDriveFilePathMap);
  }
}

await main();
