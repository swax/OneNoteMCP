import { Client } from "@microsoft/microsoft-graph-client"; // Import Client
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { registerNotebookTools } from "../functions/notebooks.js"; // Added import

const mockGraphClient = {
  api: (path: string) => {
    return {
      // Methods likely called directly after .api()
      get: async () => {
        if (path === "/me/onenote/notebooks/test-id") {
          // Response for getNotebook by ID
          return {
            id: "test-id",
            displayName: "Test Notebook",
            createdDateTime: "2024-01-01T00:00:00Z",
            lastModifiedDateTime: "2024-01-02T00:00:00Z",
            sectionsUrl:
              "https://graph.microsoft.com/v1.0/me/onenote/notebooks/test-id/sections",
          };
        }
        // Handle other GET paths if necessary, or throw error
        throw new Error(`Unhandled direct GET path: ${path}`);
      },
      post: async (data: any) => {
        if (path === "/me/onenote/notebooks") {
          // Existing create response
          return {
            id: "new-test-id",
            displayName: data.displayName, // Use provided name
            createdDateTime: "2024-01-01T00:00:00Z",
            lastModifiedDateTime: "2024-01-01T00:00:00Z",
            sectionsUrl: `https://graph.microsoft.com/v1.0/me/onenote/notebooks/new-test-id/sections`,
          };
        } else if (path.includes("/sections")) {
          // Mock section creation response (can be simple)
          return { id: "new-section-id", displayName: data.displayName };
        }
        throw new Error(`Unhandled POST path: ${path}`);
      },
      delete: async () => {
        if (path === "/me/onenote/notebooks/test-id") {
          // Existing delete response
          return undefined; // Successful delete usually returns 204 No Content
        }
        throw new Error(`Unhandled DELETE path: ${path}`);
      },
      // Keep select for listNotebooks if it uses it
      select: () => ({
        get: async () => {
          if (path === "/me/onenote/notebooks") {
            // Existing list response
            return {
              value: [
                {
                  id: "test-id",
                  displayName: "Test Notebook",
                  createdDateTime: "2024-01-01T00:00:00Z",
                  lastModifiedDateTime: "2024-01-02T00:00:00Z",
                  sectionsUrl:
                    "https://graph.microsoft.com/v1.0/me/onenote/notebooks/test-id/sections",
                },
              ],
            };
          }
          if (path === "/me/onenote/notebooks/test-id") {
            return {
              id: "test-id",
              displayName: "Test Notebook",
              createdDateTime: "2024-01-01T00:00:00Z",
              lastModifiedDateTime: "2024-01-02T00:00:00Z",
              sectionsUrl:
                "https://graph.microsoft.com/v1.0/me/onenote/notebooks/test-id/sections",
            };
          }
          // Handle other select().get() paths if necessary
          throw new Error(`Unhandled select().get() path: ${path}`);
        },
        // Add other chained methods after select if needed
      }),
      // Add other methods like patch, put if needed directly under api() return
    };
  }, // Added comma here as it's now a property assignment
};

jest.mock("@microsoft/microsoft-graph-client", () => ({
  Client: {
    // Keep the init mock as it's used by the actual implementation,
    // but the test directly uses the mockGraphClient object.
    init: () => mockGraphClient,
    // Add a static method mock if needed, or adjust based on how Client is used
    initWithMiddleware: jest.fn().mockReturnValue(mockGraphClient),
  },
}));

// Mock McpServer
const mockToolFunctions: Record<string, (args: any) => Promise<any>> = {};
const mockServer = {
  tool: jest.fn((name, description, schema, implementation) => {
    mockToolFunctions[name] = implementation;
  }),
} as unknown as McpServer; // Type assertion for mock

describe("NotebookManagement", () => {
  beforeEach(() => {
    // Clear captured functions before each test
    for (const key in mockToolFunctions) {
      delete mockToolFunctions[key];
    }
    // Reset the mock function calls
    (mockServer.tool as jest.Mock).mockClear();
    // Instantiate NotebookManagement to register tools
    registerNotebookTools(mockServer, mockGraphClient as unknown as Client); // Pass the mock client
  });

  it("should register all tools on initialization", () => {
    expect(mockServer.tool).toHaveBeenCalledWith(
      "listNotebooks",
      expect.any(String),
      {},
      expect.any(Function),
    );
    expect(mockServer.tool).toHaveBeenCalledWith(
      "getNotebook",
      expect.any(String),
      expect.any(Object),
      expect.any(Function),
    );
    expect(mockServer.tool).toHaveBeenCalledWith(
      "createNotebook",
      expect.any(String),
      expect.any(Object),
      expect.any(Function),
    );
    expect(mockServer.tool).toHaveBeenCalledWith(
      "deleteNotebook",
      expect.any(String),
      expect.any(Object),
      expect.any(Function),
    );
    // Check the keys of the captured functions
    expect(Object.keys(mockToolFunctions)).toEqual(
      expect.arrayContaining([
        "listNotebooks",
        "getNotebook",
        "createNotebook",
        "deleteNotebook",
      ]),
    );
    expect(Object.keys(mockToolFunctions).length).toBe(4); // Ensure only these 4 are registered
  });

  describe("listNotebooks tool", () => {
    it("should return notebooks in MCP format", async () => {
      const listNotebooksFn = mockToolFunctions["listNotebooks"];
      const result = await listNotebooksFn({}); // No input args for listNotebooks

      expect(result).toEqual({
        content: [
          {
            type: "resource",
            resource: {
              mimeType: "application/json",
              uri: "/me/onenote/notebooks",
              text: JSON.stringify([
                {
                  id: "test-id",
                  name: "Test Notebook",
                  createdTime: "2024-01-01T00:00:00Z",
                  lastModifiedTime: "2024-01-02T00:00:00Z",
                  sectionsUrl:
                    "https://graph.microsoft.com/v1.0/me/onenote/notebooks/test-id/sections",
                },
              ]),
            },
          },
        ],
      });
    });
  });

  describe("getNotebook tool", () => {
    it("should return a specific notebook in MCP format", async () => {
      const getNotebookFn = mockToolFunctions["getNotebook"];
      const result = await getNotebookFn({ id: "test-id" });

      expect(result).toEqual({
        content: [
          {
            type: "resource",
            resource: {
              mimeType: "application/json",
              uri: "/me/onenote/notebooks/test-id",
              text: JSON.stringify({
                id: "test-id",
                name: "Test Notebook",
                createdTime: "2024-01-01T00:00:00Z",
                lastModifiedTime: "2024-01-02T00:00:00Z",
                sectionsUrl:
                  "https://graph.microsoft.com/v1.0/me/onenote/notebooks/test-id/sections",
              }),
            },
          },
        ],
      });
    });
  });

  describe("createNotebook tool", () => {
    it("should create a new notebook and return it in MCP format", async () => {
      const createNotebookFn = mockToolFunctions["createNotebook"];
      const result = await createNotebookFn({ name: "New Test Notebook" });

      expect(result).toEqual({
        content: [
          {
            type: "resource",
            resource: {
              mimeType: "application/json",
              uri: "/me/onenote/notebooks/new-test-id",
              text: JSON.stringify({
                id: "new-test-id",
                name: "New Test Notebook",
                createdTime: "2024-01-01T00:00:00Z",
                lastModifiedTime: "2024-01-01T00:00:00Z",
                sectionsUrl:
                  "https://graph.microsoft.com/v1.0/me/onenote/notebooks/new-test-id/sections",
              }),
            },
          },
        ],
      });
    });

    it("should create a new notebook with a section if specified", async () => {
      const createNotebookFn = mockToolFunctions["createNotebook"];
      const result = await createNotebookFn({
        name: "New Test Notebook",
        sectionName: "My Section",
      });

      // We mainly check the structure, the graph mock handles the section creation side effect
      expect(result).toEqual({
        content: [
          {
            type: "resource",
            resource: {
              mimeType: "application/json",
              uri: "/me/onenote/notebooks/new-test-id",
              text: JSON.stringify({
                id: "new-test-id",
                name: "New Test Notebook",
                createdTime: "2024-01-01T00:00:00Z",
                lastModifiedTime: "2024-01-01T00:00:00Z",
                sectionsUrl:
                  "https://graph.microsoft.com/v1.0/me/onenote/notebooks/new-test-id/sections",
              }),
            },
          },
        ],
      });
    });
  });

  describe("deleteNotebook tool", () => {
    it("should delete a notebook and return success message in MCP format", async () => {
      const deleteNotebookFn = mockToolFunctions["deleteNotebook"];
      const result = await deleteNotebookFn({ id: "test-id" });

      expect(result).toEqual({
        content: [
          { type: "text", text: "Notebook test-id deleted successfully." },
        ],
      });
    });
  });
});
