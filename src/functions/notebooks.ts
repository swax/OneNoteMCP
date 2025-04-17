import { TokenCredential } from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"; // Added import
import { z } from "zod"; // Added import
import { Notebook, NotebookCreateOptions } from "../types";

export class NotebookManagement {
  private client: Client;
  private server: McpServer; // Added server property
  private credential: TokenCredential; // Added credential property

  constructor(server: McpServer, credential: TokenCredential) {
    // Updated constructor signature
    this.server = server; // Store server instance
    this.credential = credential; // Store credential instance
    this.client = Client.init({
      authProvider: async (done) => {
        try {
          const token = await this.credential.getToken(
            "https://graph.microsoft.com/.default",
          );
          done(null, token?.token || "");
        } catch (error) {
          done(error as Error, "");
        }
      },
    });

    this.registerTools(); // Call registration method
  }

  // Method to register tools
  private registerTools() {
    this.server.tool(
      "listNotebooks",
      "List all OneNote notebooks",
      {}, // No input parameters
      async (): Promise<{
        content: {
          type: "resource";
          resource: { mimeType: string; text: string };
        }[];
      }> => {
        try {
          const response = await this.client
            .api("/me/onenote/notebooks")
            .select(
              "id,displayName,createdDateTime,lastModifiedDateTime,sectionsUrl",
            )
            .get();

          const notebooks: Notebook[] = response.value.map((notebook: any) => ({
            id: notebook.id,
            name: notebook.displayName,
            createdTime: notebook.createdDateTime,
            lastModifiedTime: notebook.lastModifiedDateTime,
            sectionsUrl: notebook.sectionsUrl,
          }));
          return {
            content: [
              {
                type: "resource",
                resource: {
                  mimeType: "application/json",
                  text: JSON.stringify(notebooks),
                },
              },
            ],
          };
        } catch (error) {
          throw new Error(`Failed to list notebooks: ${error.message}`);
        }
      },
    );

    this.server.tool(
      "getNotebook",
      "Get a specific OneNote notebook by its ID",
      {
        id: z.string().describe("The ID of the notebook to retrieve"),
      },
      async ({
        id,
      }: {
        id: string;
      }): Promise<{
        content: {
          type: "resource";
          resource: { mimeType: string; text: string };
        }[];
      }> => {
        try {
          const notebook = await this.client
            .api(`/me/onenote/notebooks/${id}`)
            .select(
              "id,displayName,createdDateTime,lastModifiedDateTime,sectionsUrl",
            )
            .get();

          const result: Notebook = {
            id: notebook.id,
            name: notebook.displayName,
            createdTime: notebook.createdDateTime,
            lastModifiedTime: notebook.lastModifiedDateTime,
            sectionsUrl: notebook.sectionsUrl,
          };
          return {
            content: [
              {
                type: "resource",
                resource: {
                  mimeType: "application/json",
                  text: JSON.stringify(result),
                },
              },
            ],
          };
        } catch (error) {
          throw new Error(`Failed to get notebook ${id}: ${error.message}`);
        }
      },
    );

    this.server.tool(
      "createNotebook",
      "Create a new OneNote notebook",
      {
        name: z.string().describe("The name for the new notebook"),
        sectionName: z
          .string()
          .optional()
          .describe(
            "Optional name for an initial section within the new notebook",
          ),
      },
      async ({
        name,
        sectionName,
      }: NotebookCreateOptions): Promise<{
        content: {
          type: "resource";
          resource: { mimeType: string; text: string };
        }[];
      }> => {
        try {
          const notebook = await this.client.api("/me/onenote/notebooks").post({
            displayName: name,
          });

          if (sectionName) {
            await this.client
              .api(`/me/onenote/notebooks/${notebook.id}/sections`)
              .post({
                displayName: sectionName,
              });
          }

          // Re-fetch the notebook to get all properties including sectionsUrl if needed, or construct manually
          const createdNotebook: Notebook = {
            id: notebook.id,
            name: notebook.displayName,
            createdTime: notebook.createdDateTime,
            lastModifiedTime: notebook.lastModifiedDateTime,
            sectionsUrl: notebook.sectionsUrl, // Note: sectionsUrl might not be immediately available post-creation depending on API behavior
          };
          return {
            content: [
              {
                type: "resource",
                resource: {
                  mimeType: "application/json",
                  text: JSON.stringify(createdNotebook),
                },
              },
            ],
          };
        } catch (error) {
          throw new Error(`Failed to create notebook: ${error.message}`);
        }
      },
    );

    this.server.tool(
      "deleteNotebook",
      "Delete a OneNote notebook by its ID",
      {
        id: z.string().describe("The ID of the notebook to delete"),
      },
      async ({
        id,
      }: {
        id: string;
      }): Promise<{ content: { type: "text"; text: string }[] }> => {
        try {
          await this.client.api(`/me/onenote/notebooks/${id}`).delete();
          return {
            content: [
              { type: "text", text: `Notebook ${id} deleted successfully.` },
            ],
          };
        } catch (error) {
          throw new Error(`Failed to delete notebook ${id}: ${error.message}`);
        }
      },
    );
  }
}
