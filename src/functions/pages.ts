import { TokenCredential } from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { Page, PageCreateOptions, SearchOptions } from "../types";
import { getErrorMessage } from "../utils/error"; // Import getErrorMessage

export class PageManagement {
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
      "listPages",
      "List pages within a specific OneNote section",
      {
        sectionId: z
          .string()
          .describe("The ID of the section containing the pages"),
      },
      async ({
        sectionId,
      }: {
        sectionId: string;
      }): Promise<{
        content: {
          type: "resource";
          resource: { mimeType: string; text: string; uri: string }; // Added uri
        }[];
      }> => {
        const uri = `/me/onenote/sections/${sectionId}/pages`; // Define uri
        try {
          const response = await this.client
            .api(uri) // Use uri
            .select("id,title,createdDateTime,lastModifiedDateTime,contentUrl")
            .get();

          const pages: Page[] = response.value.map((page: any) => ({
            id: page.id,
            title: page.title,
            createdTime: page.createdDateTime,
            lastModifiedTime: page.lastModifiedDateTime,
            contentUrl: page.contentUrl,
          }));
          return {
            content: [
              {
                type: "resource",
                resource: {
                  mimeType: "application/json",
                  text: JSON.stringify(pages),
                  uri: uri, // Add uri to response
                },
              },
            ],
          };
        } catch (error) {
          throw new Error(
            `Failed to list pages in section ${sectionId}: ${getErrorMessage(
              error,
            )}`, // Use getErrorMessage
          );
        }
      },
    );

    this.server.tool(
      "searchPages",
      "Search for OneNote pages across notebooks or within a specific scope",
      {
        query: z.string().describe("The search query string"),
        notebookId: z
          .string()
          .optional()
          .describe("Optional notebook ID to limit search scope"),
        sectionId: z
          .string()
          .optional()
          .describe("Optional section ID to limit search scope"),
      },
      async ({
        query,
        notebookId,
        sectionId,
      }: SearchOptions): Promise<{
        content: {
          type: "resource";
          resource: { mimeType: string; text: string; uri: string }; // Added uri
        }[];
      }> => {
        try {
          let searchEndpoint = "/me/onenote/pages";
          if (sectionId) {
            searchEndpoint = `/me/onenote/sections/${sectionId}/pages`;
          } else if (notebookId) {
            searchEndpoint = `/me/onenote/notebooks/${notebookId}/pages`;
          }
          const uri = `${searchEndpoint}?$filter=contains(title,'${query}')`; // Define uri

          const response = await this.client
            .api(searchEndpoint) // Use base endpoint for API call
            .filter(`contains(title,'${query}')`)
            .select("id,title,createdDateTime,lastModifiedDateTime,contentUrl")
            .get();

          const pages: Page[] = response.value.map((page: any) => ({
            id: page.id,
            title: page.title,
            createdTime: page.createdDateTime,
            lastModifiedTime: page.lastModifiedDateTime,
            contentUrl: page.contentUrl,
          }));
          return {
            content: [
              {
                type: "resource",
                resource: {
                  mimeType: "application/json",
                  text: JSON.stringify(pages),
                  uri: uri, // Add uri with filter to response
                },
              },
            ],
          };
        } catch (error) {
          throw new Error(`Failed to search pages: ${getErrorMessage(error)}`); // Use getErrorMessage
        }
      },
    );

    this.server.tool(
      "createPage",
      "Create a new OneNote page within a section",
      {
        title: z.string().describe("The title for the new page"),
        content: z.string().describe("The content of the page in HTML format"),
        sectionId: z
          .string()
          .describe("The ID of the section where the page will be created"),
      },
      async ({
        title,
        content,
        sectionId,
      }: PageCreateOptions): Promise<{
        content: {
          type: "resource";
          resource: { mimeType: string; text: string; uri: string }; // Added uri
        }[];
      }> => {
        const baseUri = `/me/onenote/sections/${sectionId}/pages`; // Define base uri
        try {
          const htmlContent = `<!DOCTYPE html>
            <html>
              <head>
                <title>${title}</title>
              </head>
              <body>
                ${content}
              </body>
            </html>`;

          const page = await this.client
            .api(baseUri) // Use base uri
            .header("Content-Type", "application/xhtml+xml")
            .post(htmlContent);

          const createdPage: Page = {
            id: page.id,
            title: page.title,
            createdTime: page.createdDateTime,
            lastModifiedTime: page.lastModifiedDateTime,
            contentUrl: page.contentUrl,
          };
          const pageUri = `/me/onenote/pages/${page.id}`; // Define specific page uri
          return {
            content: [
              {
                type: "resource",
                resource: {
                  mimeType: "application/json",
                  text: JSON.stringify(createdPage),
                  uri: pageUri, // Add uri to response
                },
              },
            ],
          };
        } catch (error) {
          throw new Error(
            `Failed to create page in section ${sectionId}: ${getErrorMessage(
              error,
            )}`, // Use getErrorMessage
          );
        }
      },
    );

    this.server.tool(
      "getPageContent",
      "Get the content of a specific OneNote page",
      {
        id: z.string().describe("The ID of the page to retrieve content for"),
      },
      async ({
        id,
      }: {
        id: string;
      }): Promise<{
        content: {
          type: "resource";
          resource: { mimeType: string; text: string; uri: string }; // Added uri
        }[];
      }> => {
        const uri = `/me/onenote/pages/${id}/content`; // Define uri
        try {
          const pageMeta = await this.client
            .api(`/me/onenote/pages/${id}`)
            .select("id,title,createdDateTime,lastModifiedDateTime,contentUrl")
            .get();

          const contentStream = await this.client
            .api(uri) // Use uri
            .get();

          let pageContent = contentStream;
          if (typeof contentStream !== "string") {
            pageContent = await new Promise((resolve, reject) => {
              let data = "";
              contentStream.on("data", (chunk: any) => (data += chunk));
              contentStream.on("end", () => resolve(data));
              contentStream.on("error", (err: any) => reject(err));
            });
          }

          const resultPage: Page = {
            id: pageMeta.id,
            title: pageMeta.title,
            createdTime: pageMeta.createdDateTime,
            lastModifiedTime: pageMeta.lastModifiedDateTime,
            content: pageContent as string,
            contentUrl: pageMeta.contentUrl,
          };

          return {
            content: [
              {
                type: "resource",
                resource: {
                  mimeType: "text/html",
                  text: resultPage.content ?? "",
                  uri: uri, // Add uri to response
                },
              },
            ],
          };
        } catch (error) {
          throw new Error(
            `Failed to get content for page ${id}: ${getErrorMessage(error)}`, // Use getErrorMessage
          );
        }
      },
    );

    this.server.tool(
      "updatePageContent",
      "Update the content of an existing OneNote page",
      {
        id: z.string().describe("The ID of the page to update"),
        content: z
          .string()
          .describe(
            "The new page content in HTML format. This replaces the entire page body.",
          ),
      },
      async ({
        id,
        content,
      }: {
        id: string;
        content: string;
      }): Promise<{ content: { type: "text"; text: string }[] }> => {
        try {
          const patchPayload = [
            {
              target: "body",
              action: "replace",
              content: content,
            },
          ];

          await this.client
            .api(`/me/onenote/pages/${id}/content`)
            .header("Content-Type", "application/json")
            .patch(patchPayload);

          return {
            content: [
              {
                type: "text",
                text: `Page ${id} content updated successfully.`,
              },
            ],
          };
        } catch (error) {
          throw new Error(
            `Failed to update page ${id}: ${getErrorMessage(error)}`,
          ); // Use getErrorMessage
        }
      },
    );

    this.server.tool(
      "deletePage",
      "Delete a specific OneNote page by its ID",
      {
        id: z.string().describe("The ID of the page to delete"),
      },
      async ({
        id,
      }: {
        id: string;
      }): Promise<{ content: { type: "text"; text: string }[] }> => {
        try {
          await this.client.api(`/me/onenote/pages/${id}`).delete();
          return {
            content: [
              { type: "text", text: `Page ${id} deleted successfully.` },
            ],
          };
        } catch (error) {
          throw new Error(
            `Failed to delete page ${id}: ${getErrorMessage(error)}`,
          ); // Use getErrorMessage
        }
      },
    );
  }
}
