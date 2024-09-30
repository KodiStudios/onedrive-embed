import { Client } from "@microsoft/microsoft-graph-client";
import minimist from "minimist";

function Main(): void {
  const argv = minimist(process.argv.slice(2));
  if (!argv.token) {
    console.log("Usage: ");
    console.log(
      "node --experimental-strip-types index.ts --token {token_value_from_aka.ms/ge}"
    );
    return;
  }

  const graphClient = Client.init({
    defaultVersion: "v1.0",
    debugLogging: true,
    authProvider: (done) => {
      const errorMessage = "error throw by the authentication handler";
      done(errorMessage, argv.token);
    },
  });

  // Profile Api
  graphClient
    .api("/me")
    .select("displayName")
    .get()
    .then((res) => {
      console.log(res);
    })
    .catch((err) => {
      console.log(err);
    });
}

Main();
