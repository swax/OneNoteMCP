import { Client } from "@microsoft/microsoft-graph-client";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import z from "zod";
import { Notebook, NotebookCreateOptions } from "../types";
import { getErrorMessage } from "../utils/error";

export function registerNotebookTools(
  mcpServer: McpServer,
  azureClient: Client,
) {
  mcpServer.tool(
    "listNotebooks",
    "List all OneNote notebooks",
    {}, // No input parameters
    async (args, extra): Promise<{
      content: {
        type: "resource";
        resource: { mimeType: string; text: string; uri: string };
      }[];
    }> => {
      const uri = "/me/onenote/notebooks";
      try {
        const response = await azureClient
          .api(uri)
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
                uri,
              },
            },
          ],
        };
      } catch (error) {
        throw new Error(`Failed to list notebooks: ${getErrorMessage(error)}`);
      }
    },
  );

  mcpServer.tool(
    "getNotebook",
    "Get a specific OneNote notebook by its ID",
    {
      id: z.string().describe("The ID of the notebook to retrieve"),
    },
    async (args, extra): Promise<{
      content: {
        type: "resource";
        resource: { mimeType: string; text: string; uri: string };
      }[];
    }> => {
      const { id } = args;
      const uri = `/me/onenote/notebooks/${id}`;
      try {
        const notebook = await azureClient
          .api(uri)
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
                uri,
              },
            },
          ],
        };
      } catch (error) {
        throw new Error(
          `Failed to get notebook ${id}: ${getErrorMessage(error)}`,
        );
      }
    },
  );

  mcpServer.tool(
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
    async (args, extra): Promise<{
      content: {
        type: "resource";
        resource: { mimeType: string; text: string; uri: string };
      }[];
    }> => {
      const { name, sectionName } = args;
      const baseUri = "/me/onenote/notebooks";
      try {
        const notebook = await azureClient.api(baseUri).post({
          displayName: name,
        });

        if (sectionName) {
          await azureClient
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
        const notebookUri = `/me/onenote/notebooks/${notebook.id}`;
        return {
          content: [
            {
              type: "resource",
              resource: {
                mimeType: "application/json",
                text: JSON.stringify(createdNotebook),
                uri: notebookUri,
              },
            },
          ],
        };
      } catch (error) {
        throw new Error(`Failed to create notebook: ${getErrorMessage(error)}`);
      }
    },
  );

  mcpServer.tool(
    "deleteNotebook",
    "Delete a OneNote notebook by its ID",
    {
      id: z.string().describe("The ID of the notebook to delete"),
    },
    async (args, extra): Promise<{ content: { type: "text"; text: string }[] }> => {
      const { id } = args;
      const uri = `/me/onenote/notebooks/${id}`;
      try {
        await azureClient.api(uri).delete();
        return {
          content: [
            { type: "text", text: `Notebook ${id} deleted successfully.` },
          ],
        };
      } catch (error) {
        throw new Error(
          `Failed to delete notebook ${id}: ${getErrorMessage(error)}`,
        );
      }
    },
  );
}
