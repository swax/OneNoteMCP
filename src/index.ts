import {
  AuthenticationRecord,
  ClientSecretCredential,
  DeviceCodeCredential,
  useIdentityPlugin,
} from "@azure/identity";
import { cachePersistencePlugin } from "@azure/identity-cache-persistence";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import fs from "fs";
import { NotebookManagement } from "./functions/notebooks";
import { PageManagement } from "./functions/pages";
import { SectionManagement } from "./functions/sections";

require("dotenv").config(); // Load environment variables from .env file

export class OneNoteMCPServer extends McpServer {
  private credential: ClientSecretCredential;
  private deviceCodeCredential: DeviceCodeCredential | undefined = undefined;

  constructor() {
    super({
      name: "AzureOneNoteMCP",
      version: "1.0.0",
    });

    const clientId = process.env.AZURE_CLIENT_ID;
    const clientSecret = process.env.AZURE_CLIENT_SECRET;

    if (!clientId || !clientSecret) {
      throw new Error(
        "Azure credentials must be provided via environment variables",
      );
    }

    useIdentityPlugin(cachePersistencePlugin);

    const authenticationRecordPath = ".cache/authentication-record.json";
    let authenticationRecord: AuthenticationRecord | undefined = undefined;

    // If we have an existing record, deserialize it.
    if (fs.existsSync(authenticationRecordPath)) {
      // authenticationRecord = AuthenticationRecord.deserialize(new FileInputStream(authenticationRecordPath));
      authenticationRecord = JSON.parse(
        fs.readFileSync(authenticationRecordPath, "utf-8"),
      );
      console.log("Loaded cached authentication record");
    }

    const tenantId = "consumers";

    this.credential = new ClientSecretCredential(
      tenantId,
      clientId,
      clientSecret,
    );

    this.deviceCodeCredential = new DeviceCodeCredential({
      clientId,
      tenantId,
      tokenCachePersistenceOptions: {
        enabled: true,
      },
      authenticationRecord,
      /*userPromptCallback: (info: DeviceCodeInfo) => {
        // Display the device code message to
        // the user. This tells them
        // where to go to sign in and provides the
        // code to use.
        console.log(info.message);
      },*/
    });

    const scope = "https://graph.microsoft.com/.default";

    this.deviceCodeCredential.authenticate(scope).then((token) => {
      // Save the authentication record to a file
      console.log("Authenticated successfully");
      fs.writeFileSync(authenticationRecordPath, JSON.stringify(token));
    });

    /*this.deviceCodeCredential
      .getToken(scope)
      .then((token) => console.log("GetToken: ", token));*/

    new NotebookManagement(this, this.deviceCodeCredential);
    new PageManagement(this, this.credential);
    new SectionManagement(this, this.credential);
  }

  async run() {
    const transport = new StdioServerTransport();
    await this.server.connect(transport);
    console.error("OneNote MCP server running on stdio");
  }
}

const server = new OneNoteMCPServer();
server.run().catch(console.error);
