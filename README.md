Forked from [ZubeidHendricks/azure-onenote-mcp-server](https://github.com/ZubeidHendricks/azure-onenote-mcp-server) see that page for the original README.md

## About

An MCP for OneNote. Forked and revised extensively, now works with personal notebooks.
Updated to the latest MCP API. Supports caching credentials.

## Improvements

### Functionality

- Fixed `getPageContent` functionality by implementing ReadableStream
- Successfully implemented API calls against personal notebooks
- Updated index and tools to use the latest MCP version

### Code Architecture

- Simplified codebase by removing classes that weren't necessary
- Many unnecessary files removed

### Authentication & Performance

- Implemented disk caching for authentication credentials
- Updated package dependencies
- Added dotenv support for direct testing capabilities

## Azure App Registration

1. Go to Azure Portal and navigate to App registrations
2. Create a new registration for OneNoteMCP
3. Add Microsoft Graph API permissions:
   - Notes.Read
   - Notes.Read.All
   - (optionally add Notes.Write permissions, but there is a risk of losing your notes)
4. Create a client secret
5. Copy the client ID/secret for configuration
6. In the manifest set `signInAudience` to `AzureADandPersonalMicrosoftAccount`

## MCP Server setup

Put the client ID/secret n your `.env` file locally as well as in Claude desktop.

The `AUTH_CACHE_DIR` should be the directory where you want the generated API access token to be stored.

```
AZURE_CLIENT_ID=\
AZURE_CLIENT_SECRET=
AUTH_CACHE_DIR="C:\\git\\azure-onenote-mcp-server\\.cache"
```

Run server manually first to generate the credentials that will be cached and used by
Claude desktop when it runs the server later on.

```
npm install
npm build
npm run
```

You can test the API works by running this command directly once the server starts up and logs in.

`{"jsonrpc": "2.0", "id": 3, "method": "tools/call", "params": {"name": "listNotebooks", "arguments": {}}}`

## Claude MCP config

```
{
  "mcpServers": {
    "onenote": {
      "command": "node",
      "args": ["C:\\git\\azure-onenote-mcp-server\\dist\\index.js"],
      "env": {
        "AZURE_CLIENT_ID": "...",
        "AZURE_CLIENT_SECRET": "...",
        "AUTH_CACHE_DIR": "C:\\git\\azure-onenote-mcp-server\\.cache"
      }
    }
  }
}
```
