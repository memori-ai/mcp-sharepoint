#!/usr/bin/env node

import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";
import { getFolders, createFolder, deleteFolder, getFolderTree } from "./services/folderService.js";
import { getDocuments, getDocumentContent, uploadDocument, updateDocumentContent, deleteDocument, searchDocumentsByKeywords } from "./services/fileService.js";

// Crea il server
const server = new Server(
  {
    name: "aisuru-mcp-server-sharepoint",
    version: "1.0.0",
  },
  {
    capabilities: {
      tools: {},
    },
  }
);

const authParams = {
  type: "object",
  optional: true,
  description: "Microsoft Entra ID app authentication parameters",
  properties: {
    tenantId: {
      type: "string",
      description: "The directory (tenant) ID",
    },
    clientId: {
      type: "string",
      description: "The application (client) ID",
    },
    clientSecret: {
      type: "string",
      description: "The client secret",
    },
  },
}

const siteDriveParams = {
  type: "object",
  description: "SharePoint site and drive identifiers",
  properties: {
    siteId: {
      type: "string",
      description: "The ID of the SharePoint site",
    },
    driveId: {
      type: "string",
      description: "The ID of the drive within the SharePoint site",
    },
  },
};

// Lista dei tool disponibili
server.setRequestHandler(ListToolsRequestSchema, async () => {
  return {
    tools: [
      {
        name: "getFolders",
        description: "Retrieve a list of folders from the specified path in SharePoint",
        inputSchema: {
          type: "object",
          properties: {
            auth: authParams,
            siteDrive: siteDriveParams,
            path: { type: "string", description: "The path in SharePoint to retrieve folders from" },
          },
          required: ["auth", "siteDrive", "path"],
        },
      },
      {
        name: "createFolder",
        description: "Create a new folder in SharePoint at the specified path",
        inputSchema: {
          type: "object",
          properties: {
            auth: authParams,
            siteDrive: siteDriveParams,
            path: { type: "string", description: "The parent path where the folder will be created" },
            folderName: { type: "string", description: "The name of the new folder to create" },
          },
          required: ["auth", "siteDrive", "path", "folderName"],
        },
      },
      {
        name: "deleteFolder",
        description: "Delete an empty folder in SharePoint at the specified path",
        inputSchema: {
          type: "object",
          properties: {
            auth: authParams,
            siteDrive: siteDriveParams,
            path: { type: "string", description: "The path of the folder to delete" },
          },
          required: ["auth", "siteDrive", "path"],
        },
      },
      {
        name: "getFolderTree",
        description: "Get a tree view of the folder structure in SharePoint",
        inputSchema: {
          type: "object",
          properties: {
            auth: authParams,
            siteDrive: siteDriveParams,
            path: { type: "string", description: "The starting path (default: 'root')" },
            maxDepth: { type: "number", description: "Maximum depth to traverse (default: 10)" },
          },
          required: ["auth", "siteDrive"],
        },
      },
      {
        name: "getDocuments",
        description: "List all documents and their metadata in a specified path in SharePoint",
        inputSchema: {
          type: "object",
          properties: {
            auth: authParams,
            siteDrive: siteDriveParams,
            path: { type: "string", description: "The path in SharePoint to retrieve documents from" },
          },
          required: ["auth", "siteDrive", "path"],
        },
      },
      {
        name: "getDocumentContent",
        description: "Get the content of a document in SharePoint, supporting multiple formats (PDF, Word, Excel)",
        inputSchema: {
          type: "object",
          properties: {
            auth: authParams,
            siteDrive: siteDriveParams,
            filePath: { type: "string", description: "The path to the file (e.g., 'Cartella_1/file.docx')" },
          },
          required: ["auth", "siteDrive", "filePath"],
        },
      },
      {
        name: "uploadDocument",
        description: "Upload a document to a specified path in SharePoint",
        inputSchema: {
          type: "object",
          properties: {
            auth: authParams,
            siteDrive: siteDriveParams,
            filePath: { type: "string", description: "The path where the file will be uploaded (e.g., 'Cartella_1/file.docx')" },
            content: { type: "string", description: "The string or base64-encoded content of the file to upload" },
            contentType: { type: "string", description: "The MIME type of the content (e.g., 'application/pdf')" },
            overwrite: { type: "boolean", description: "Whether to overwrite existing files" },
          },
          required: ["auth", "siteDrive", "filePath", "content"],
        },
      },
      {
        name: "updateDocumentContent",
        description: "Update the content of an existing document in SharePoint, Replaces the entire content.",
        inputSchema: {
          type: "object",
          properties: {
            auth: authParams,
            siteDrive: siteDriveParams,
            filePath: { type: "string", description: "The path to the existing file to update (e.g., 'Cartella_1/file.docx')" },
            content: { type: "string", description: "The new string or base64-encoded content of the file" },
            contentType: { type: "string", description: "The MIME type of the content (e.g., 'application/pdf')" },
          },
          required: ["auth", "siteDrive", "filePath", "content"],
        },
      },
      {
        name: "deleteDocument",
        description: "Delete a document in SharePoint at the specified path",
        inputSchema: {
          type: "object",
          properties: {
            auth: authParams,
            siteDrive: siteDriveParams,
            filePath: { type: "string", description: "The path to the file to delete (e.g., 'Cartella_1/file.docx')" },
          },
          required: ["auth", "siteDrive", "filePath"],
        },
      },
      {
        name: "searchDocumentsByKeywords",
        description: "Search for documents in SharePoint containing specific keywords in the given attribute",
        inputSchema: {
          type: "object",
          properties: {
            auth: authParams,
            siteDrive: siteDriveParams,
            listId: { type: "string", description: "The ID of the SharePoint list to search in" },
            keywords: { type: "array", items: { type: "string" }, description: "List of keywords to search for" },
            attributeName: { type: "string", description: "The document attribute to search in (e.g., 'name', 'content')" },
          },
          required: ["auth", "siteDrive", "listId", "keywords", "attributeName"],
        },
      }
    ],
  };
});

// Gestione delle chiamate ai tool
server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params;

  switch (name) {
    case "getFolders":
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(await getFolders(args.auth, args.siteDrive, args.path)),
          },
        ],
      };

    case "createFolder":
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(await createFolder(args.auth, args.siteDrive, args.path, args.folderName)),
          },
        ],
      };

    case "deleteFolder":
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(await deleteFolder(args.auth, args.siteDrive, args.path)),
          },
        ],
      };

    case "getFolderTree":
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(await getFolderTree(args.auth, args.siteDrive, args.path, args.maxDepth)),
          },
        ],
      };

    case "getDocuments":
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(await getDocuments(args.auth, args.siteDrive, args.path)),
          },
        ],
      };

    case "getDocumentContent":
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(await getDocumentContent(args.auth, args.siteDrive, args.filePath)),
          },
        ],
      };

    case "uploadDocument":
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(await uploadDocument(args.auth, args.siteDrive, args.filePath, args.content, args.contentType, args.overwrite)),
          },
        ],
      };

    case "updateDocumentContent":
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(await updateDocumentContent(args.auth, args.siteDrive, args.filePath, args.content, args.contentType)),
          },
        ],
      };

    case "deleteDocument":
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(await deleteDocument(args.auth, args.siteDrive, args.filePath)),
          },
        ],
      };

    case "searchDocumentsByKeywords":
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(await searchDocumentsByKeywords(args.auth, args.siteDrive, args.listId, args.keywords, args.attributeName)),
          },
        ],
      };

    default:
      throw new Error(`Tool sconosciuto: ${name}`);
  }
});

// Avvia il server
async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error("Server MCP avviato");
}

main().catch((error) => {
  console.error("Errore:", error);
  process.exit(1);
});