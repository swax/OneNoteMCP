import { TokenCredential } from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { Section, SectionCreateOptions } from "../types";

export class SectionManagement {
  private client: Client;
  private server: McpServer;
  private credential: TokenCredential;

  constructor(server: McpServer, credential: TokenCredential) {
    this.server = server;
    this.credential = credential;
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
    this.registerTools();
  }

  private registerTools() {
    this.server.tool(
      "listSections",
      "List sections within a specific OneNote notebook",
      {
        notebookId: z
          .string()
          .describe("The ID of the notebook containing the sections"),
      },
      async ({
        notebookId,
      }: {
        notebookId: string;
      }): Promise<{
        content: {
          type: "resource";
          resource: { mimeType: string; text: string };
        }[];
      }> => {
        try {
          const response = await this.client
            .api(`/me/onenote/notebooks/${notebookId}/sections`)
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
                },
              },
            ],
          };
        } catch (error) {
          throw new Error(
            `Failed to list sections in notebook ${notebookId}: ${error.message}`,
          );
        }
      },
    );

    this.server.tool(
      "createSection",
      "Create a new section within a OneNote notebook",
      {
        name: z.string().describe("The name for the new section"),
        notebookId: z
          .string()
          .describe("The ID of the notebook where the section will be created"),
      },
      async ({
        name,
        notebookId,
      }: SectionCreateOptions): Promise<{
        content: {
          type: "resource";
          resource: { mimeType: string; text: string };
        }[];
      }> => {
        try {
          const section = await this.client
            .api(`/me/onenote/notebooks/${notebookId}/sections`)
            .post({
              displayName: name,
            });

          const createdSection: Section = {
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
                  text: JSON.stringify(createdSection),
                },
              },
            ],
          };
        } catch (error) {
          throw new Error(
            `Failed to create section in notebook ${notebookId}: ${error.message}`,
          );
        }
      },
    );

    this.server.tool(
      "getSection",
      "Get details of a specific OneNote section by its ID",
      {
        id: z.string().describe("The ID of the section to retrieve"),
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
          const section = await this.client
            .api(`/me/onenote/sections/${id}`)
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
                },
              },
            ],
          };
        } catch (error) {
          throw new Error(`Failed to get section ${id}: ${error.message}`);
        }
      },
    );

    this.server.tool(
      "deleteSection",
      "Delete a specific OneNote section by its ID",
      {
        id: z.string().describe("The ID of the section to delete"),
      },
      async ({
        id,
      }: {
        id: string;
      }): Promise<{ content: { type: "text"; text: string }[] }> => {
        try {
          await this.client.api(`/me/onenote/sections/${id}`).delete();
          return {
            content: [
              { type: "text", text: `Section ${id} deleted successfully.` },
            ],
          };
        } catch (error) {
          throw new Error(`Failed to delete section ${id}: ${error.message}`);
        }
      },
    );
  }
}
