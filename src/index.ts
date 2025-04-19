import {
  ClientSecretCredential,
  DeviceCodeCredential,
  DeviceCodeInfo,
} from "@azure/identity";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
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

    const tenantId = process.env.AZURE_TENANT_ID;
    const clientId = process.env.AZURE_CLIENT_ID;
    const clientSecret = process.env.AZURE_CLIENT_SECRET;

    if (!tenantId || !clientId || !clientSecret) {
      throw new Error(
        "Azure credentials must be provided via environment variables",
      );
    }

    this.credential = new ClientSecretCredential(
      tenantId,
      clientId,
      clientSecret,
    );

    this.deviceCodeCredential = new DeviceCodeCredential({
      clientId: clientId,
      tenantId: "consumers",
      userPromptCallback: (info: DeviceCodeInfo) => {
        // Display the device code message to
        // the user. This tells them
        // where to go to sign in and provides the
        // code to use.
        console.log(info.message);
      },
    });

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
