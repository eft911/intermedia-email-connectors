import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { createMcpExpressApp } from "@modelcontextprotocol/sdk/server/express.js";
import { z } from "zod/v4";
import { EwsClient, loadConfig, PROCESSED_CATEGORY } from "./ews-client.js";

const config = loadConfig();
const client = new EwsClient(config.ews);

function result(value) {
  return {
    content: [{ type: "text", text: JSON.stringify(value, null, 2) }],
    structuredContent: value,
  };
}

function toolError(error) {
  return {
    isError: true,
    content: [{ type: "text", text: error instanceof Error ? error.message : "Unknown connector error" }],
  };
}

function createServer() {
  const server = new McpServer({
    name: "intermedia-exchange-connector",
    version: "0.1.0",
  });

  server.registerTool("intermedia_health_check", {
    title: "Check customer-service mailbox",
    description: "Read-only. Verify EWS authentication and access to the configured customer-service mailbox.",
    inputSchema: {},
  }, async () => {
    try { return result(await client.healthCheck()); }
    catch (error) { return toolError(error); }
  });

  server.registerTool("list_customer_messages", {
    title: "List customer-service messages",
    description: `Read-only. List recent messages in the configured customer-service inbox. Messages categorized ${PROCESSED_CATEGORY} are excluded by default to prevent duplicate drafts.`,
    inputSchema: {
      unread_only: z.boolean().default(true).describe("Return only unread messages."),
      since_hours: z.number().int().min(1).max(720).default(72).describe("Only return messages received within this many hours."),
      max_results: z.number().int().min(1).max(50).default(20).describe("Maximum messages to return."),
      include_already_drafted: z.boolean().default(false).describe(`Include messages already categorized ${PROCESSED_CATEGORY}.`),
    },
  }, async ({ unread_only, since_hours, max_results, include_already_drafted }) => {
    try {
      return result(await client.listMessages({
        unreadOnly: unread_only,
        sinceHours: since_hours,
        maxResults: max_results,
        includeProcessed: include_already_drafted,
      }));
    } catch (error) { return toolError(error); }
  });

  server.registerTool("get_customer_message", {
    title: "Read a customer-service message",
    description: "Read-only. Retrieve the full plain-text body and metadata for one customer-service message.",
    inputSchema: {
      message_id: z.string().min(1).describe("Opaque EWS message ID returned by list_customer_messages."),
    },
  }, async ({ message_id }) => {
    try { return result(await client.getMessage(message_id)); }
    catch (error) { return toolError(error); }
  });

  server.registerTool("create_customer_reply_draft", {
    title: "Save customer reply draft",
    description: `Create a plain-text reply draft in the configured customer-service Drafts folder. This never sends email. After saving, it adds the ${PROCESSED_CATEGORY} category to the source message to prevent duplicate drafts.`,
    inputSchema: {
      message_id: z.string().min(1).describe("Opaque EWS message ID returned by list_customer_messages."),
      text_content: z.string().min(1).max(20_000).describe("Plain-text reply body. Do not include private notes."),
    },
  }, async ({ message_id, text_content }) => {
    try { return result(await client.createReplyDraft({ messageId: message_id, textContent: text_content })); }
    catch (error) { return toolError(error); }
  });

  return server;
}

const app = createMcpExpressApp();
const mcpPath = `/mcp/${config.mcpPathSecret}`;

app.get("/health", (_req, res) => {
  res.json({ ok: true, service: "intermedia-exchange-connector" });
});

app.all("/mcp", (_req, res) => res.status(404).json({ error: "Not found" }));
app.post(mcpPath, async (req, res) => {
  const server = createServer();
  const transport = new StreamableHTTPServerTransport({ sessionIdGenerator: undefined });
  try {
    await server.connect(transport);
    await transport.handleRequest(req, res, req.body);
  } catch (error) {
    console.error("MCP request failed", error instanceof Error ? error.message : "Unknown error");
    if (!res.headersSent) {
      res.status(500).json({ jsonrpc: "2.0", error: { code: -32603, message: "Internal server error" }, id: null });
    }
  } finally {
    res.on("close", () => {
      transport.close();
      server.close();
    });
  }
});

app.get(mcpPath, (_req, res) => res.status(405).json({ error: "Method not allowed" }));
app.delete(mcpPath, (_req, res) => res.status(405).json({ error: "Method not allowed" }));

app.listen(config.port, (error) => {
  if (error) {
    console.error("Failed to start connector", error);
    process.exit(1);
  }
  console.log(`Intermedia Exchange connector listening on port ${config.port}`);
});
