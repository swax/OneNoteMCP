import { ClientSecretCredential } from "@azure/identity";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { NotebookManagement } from "./functions/notebooks";
import { PageManagement } from "./functions/pages";
import { SectionManagement } from "./functions/sections";

export class OneNoteMCPServer extends McpServer {
  private credential: ClientSecretCredential;

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

    new NotebookManagement(this, this.credential);
    new PageManagement(this, this.credential);
    // Instantiate SectionManagement, passing the server instance (this)
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
