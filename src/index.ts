import { DeviceCodeCredential, useIdentityPlugin } from "@azure/identity";
import { cachePersistencePlugin } from "@azure/identity-cache-persistence";
import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { registerNotebookTools } from "./functions/notebooks";
import { registerPageTools } from "./functions/pages";
import { readAuthRecordCache, writeAuthRecordCache } from "./utils/authCache";

require("dotenv").config(); // Load environment variables from .env file

const mcpServer = new McpServer({
  name: "AzureOneNoteMCP",
  version: "1.0.0",
});

// Set up the Azure credentials
const clientId = process.env.AZURE_CLIENT_ID;
const clientSecret = process.env.AZURE_CLIENT_SECRET;

if (!clientId || !clientSecret) {
  throw new Error(
    "Azure credentials must be provided via environment variables",
  );
}

useIdentityPlugin(cachePersistencePlugin);

const deviceCodeCredential = new DeviceCodeCredential({
  clientId,
  tenantId: "consumers",
  tokenCachePersistenceOptions: {
    enabled: true,
  },
  authenticationRecord: readAuthRecordCache(),
});

const scope = "https://graph.microsoft.com/.default";

// Authenticate and cache the token
deviceCodeCredential
  .authenticate(scope)
  .then((record) => {
    if (record) {
      writeAuthRecordCache(record);
    }
  })
  .catch((error) => {
    console.error("Authentication failed:", error);
    // Handle authentication failure, e.g., exit the process or notify the user
    process.exit(1);
  });

const authProvider = new TokenCredentialAuthenticationProvider(
  deviceCodeCredential,
  {
    scopes: [scope], // Ensure scopes are passed as an array
  },
);

const azureClient = Client.initWithMiddleware({
  authProvider,
});

// Instantiate management classes with the McpServer instance and the client
registerNotebookTools(mcpServer, azureClient);
registerPageTools(mcpServer, azureClient);
//new SectionManagement(mcpServer, client);

// Run the server
async function runServer() {
  const transport = new StdioServerTransport();
  await mcpServer.server.connect(transport);
  console.error("OneNote MCP server running on stdio");
}

runServer().catch(console.error);
