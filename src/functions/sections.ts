import { Client } from "@microsoft/microsoft-graph-client";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import z from "zod";
import { Section, SectionCreateOptions } from "../types";
import { getErrorMessage } from "../utils/error";

export function registerSectionTools(mcpServer: McpServer, azureClient: Client) {
  mcpServer.tool(
    "listSections",
    "List sections within a specific OneNote notebook",
    {
      notebookId: z
        .string()
        .describe("The ID of the notebook containing the sections"),
    },
    async (args, extra): Promise<{
      content: {
        type: "resource";
        resource: { mimeType: string; text: string; uri: string };
      }[];
    }> => {
      const { notebookId } = args;
      const uri = `/me/onenote/notebooks/${notebookId}/sections`;
      try {
        const response = await azureClient
          .api(uri)
          .select(
            "id,displayName,createdDateTime,lastModifiedDateTime,pagesUrl",
          )
          .get();

        const sections: Section[] = response.value.map((section: any) => ({
          id: section.id,
          name: section.displayName,
          createdTime: section.createdDateTime,
          lastModifiedTime: section.lastModifiedDateTime,
          pagesUrl: section.pagesUrl,
        }));
        return {
          content: [
            {
              type: "resource",
              resource: {
                mimeType: "application/json",
                text: JSON.stringify(sections),
                uri,
              },
            },
          ],
        };
      } catch (error) {
        throw new Error(
          `Failed to list sections in notebook ${notebookId}: ${getErrorMessage(
            error,
          )}`,
        );
      }
    },
  );

  mcpServer.tool(
    "createSection",
    "Create a new section within a OneNote notebook",
    {
      name: z.string().describe("The name for the new section"),
      notebookId: z
        .string()
        .describe("The ID of the notebook where the section will be created"),
    },
    async (args, extra): Promise<{
      content: {
        type: "resource";
        resource: { mimeType: string; text: string; uri: string };
      }[];
    }> => {
      const { name, notebookId } = args;
      const baseUri = `/me/onenote/notebooks/${notebookId}/sections`;
      try {
        const section = await azureClient.api(baseUri).post({
          displayName: name,
        });

        const createdSection: Section = {
          id: section.id,
          name: section.displayName,
          createdTime: section.createdDateTime,
          lastModifiedTime: section.lastModifiedDateTime,
          pagesUrl: section.pagesUrl,
        };
        const sectionUri = `/me/onenote/sections/${section.id}`;
        return {
          content: [
            {
              type: "resource",
              resource: {
                mimeType: "application/json",
                text: JSON.stringify(createdSection),
                uri: sectionUri,
              },
            },
          ],
        };
      } catch (error) {
        throw new Error(
          `Failed to create section in notebook ${notebookId}: ${getErrorMessage(
            error,
          )}`,
        );
      }
    },
  );

  mcpServer.tool(
    "getSection",
    "Get details of a specific OneNote section by its ID",
    {
      id: z.string().describe("The ID of the section to retrieve"),
    },
    async (args, extra): Promise<{
      content: {
        type: "resource";
        resource: { mimeType: string; text: string; uri: string };
      }[];
    }> => {
      const { id } = args;
      const uri = `/me/onenote/sections/${id}`;
      try {
        const section = await azureClient
          .api(uri)
          .select(
            "id,displayName,createdDateTime,lastModifiedDateTime,pagesUrl",
          )
          .get();

        const resultSection: Section = {
          id: section.id,
          name: section.displayName,
          createdTime: section.createdDateTime,
          lastModifiedTime: section.lastModifiedDateTime,
          pagesUrl: section.pagesUrl,
        };
        return {
          content: [
            {
              type: "resource",
              resource: {
                mimeType: "application/json",
                text: JSON.stringify(resultSection),
                uri,
              },
            },
          ],
        };
      } catch (error) {
        throw new Error(
          `Failed to get section ${id}: ${getErrorMessage(error)}`,
        );
      }
    },
  );

  mcpServer.tool(
    "deleteSection",
    "Delete a specific OneNote section by its ID",
    {
      id: z.string().describe("The ID of the section to delete"),
    },
    async (args, extra): Promise<{ content: { type: "text"; text: string }[] }> => {
      const { id } = args;
      try {
        await azureClient.api(`/me/onenote/sections/${id}`).delete();
        return {
          content: [
            { type: "text", text: `Section ${id} deleted successfully.` },
          ],
        };
      } catch (error) {
        throw new Error(
          `Failed to delete section ${id}: ${getErrorMessage(error)}`,
        );
      }
    },
  );
}
