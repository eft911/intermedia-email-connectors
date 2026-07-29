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

function attachmentResult(value) {
  const { image_base64: imageBase64, ...metadata } = value;
  const content = [{ type: "text", text: JSON.stringify(metadata, null, 2) }];
  if (imageBase64) content.push({ type: "image", data: imageBase64, mimeType: value.content_type });
  return { content, structuredContent: metadata };
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
    version: "0.5.0",
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

  server.registerTool("search_elton_mail", {
    title: "Search Elton's mail",
    description: "Read-only. Search 1, 2, or 3 years of the configured etucker@metooshoes.com mailbox across Inbox, Sent Items, and Archive. Matches both message subjects and plain-text bodies and returns an opaque cursor when more results are available.",
    inputSchema: {
      query: z.string().min(2).max(200).describe("Text or phrase to match in the subject or message body."),
      lookback_years: z.number().int().min(1).max(3).default(1).describe("Number of calendar years to search: 1, 2, or 3."),
      page_size: z.number().int().min(1).max(50).default(20).describe("Maximum results to return on this page."),
      cursor: z.string().min(1).nullable().default(null).describe("Opaque next_cursor from the previous page. Omit for the first page."),
    },
  }, async ({ query, lookback_years, page_size, cursor }) => {
    try { return result(await client.searchMessages({ query, lookbackYears: lookback_years, pageSize: page_size, cursor })); }
    catch (error) { return toolError(error); }
  });

  server.registerTool("list_elton_attachments", {
    title: "List Elton email attachments",
    description: "Read-only. List every file or attached-email attachment on a specified etucker@metooshoes.com message, including name, type, size, and reading capability.",
    inputSchema: {
      message_id: z.string().min(1).describe("Opaque EWS message ID returned by a mail search or message list."),
    },
  }, async ({ message_id }) => {
    try { return result(await client.listAttachments(message_id)); }
    catch (error) { return toolError(error); }
  });

  server.registerTool("read_elton_attachment", {
    title: "Read Elton email attachment",
    description: "Read-only. Read one attachment from an etucker@metooshoes.com message. Extracts text from common PDF, Word, Excel, PowerPoint, OpenDocument, RTF, CSV, HTML, Markdown, and text/data formats; returns common images visually; and expands attached emails. Maximum size is 10 MB.",
    inputSchema: {
      message_id: z.string().min(1).describe("Opaque EWS message ID containing the attachment."),
      attachment_id: z.string().min(1).describe("Opaque EWS attachment ID returned by list_elton_attachments."),
    },
  }, async ({ message_id, attachment_id }) => {
    try { return attachmentResult(await client.readAttachment({ messageId: message_id, attachmentId: attachment_id })); }
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

  server.registerTool("create_elton_email_draft", {
    title: "Save Elton outbound email draft",
    description: "Create a new plain-text outbound email draft in the configured etucker@metooshoes.com Drafts folder. This action only saves a draft and never sends email.",
    inputSchema: {
      to: z.array(z.email()).min(1).max(50).describe("Recipient email addresses. At least one is required."),
      cc: z.array(z.email()).max(50).default([]).describe("Optional CC recipient email addresses."),
      bcc: z.array(z.email()).max(50).default([]).describe("Optional BCC recipient email addresses."),
      subject: z.string().min(1).max(255).describe("Email subject."),
      text_content: z.string().min(1).max(20_000).describe("Plain-text email body. Do not include private notes."),
    },
  }, async ({ to, cc, bcc, subject, text_content }) => {
    try {
      return result(await client.createEmailDraft({
        to,
        cc,
        bcc,
        subject,
        textContent: text_content,
      }));
    } catch (error) { return toolError(error); }
  });

  server.registerTool("finalize_elton_outreach_drafts", {
    title: "Update and send reviewed Elton outreach drafts",
    description: "Send reviewed outreach drafts from etucker@metooshoes.com. Before any write, this action requires an exact subject match, exact expected draft count, a trailing signature match, and at least one recipient on every draft. It replaces only that trailing signature, sends the drafts, and saves copies in Sent Items.",
    inputSchema: {
      subject: z.string().min(1).max(255).describe("Exact subject shared by the reviewed outreach drafts."),
      expected_count: z.number().int().min(1).max(250).describe("Exact number of matching drafts required. If the count differs, nothing is changed or sent."),
      old_signature: z.string().min(1).max(200).describe("Exact trailing signature currently present on every matching draft, including any leading newline."),
      new_signature: z.string().min(1).max(200).describe("Replacement trailing signature, including any leading newline."),
      confirm_send: z.literal(true).describe("Must be true to confirm that the reviewed drafts should be sent now."),
    },
  }, async ({ subject, expected_count, old_signature, new_signature }) => {
    try {
      return result(await client.finalizeOutreachDrafts({
        subject,
        expectedCount: expected_count,
        oldSignature: old_signature,
        newSignature: new_signature,
      }));
    } catch (error) { return toolError(error); }
  });

  return server;
}

// Render reaches the service through its own hostname. The SDK's localhost
// default rejects those health-check Host headers, so bind as a hosted service.
// MCP access remains protected by the high-entropy secret path below.
const app = createMcpExpressApp({ host: "0.0.0.0" });
const mcpPath = `/mcp/${config.mcpPathSecret}`;

app.get("/health", (_req, res) => {
  res.json({
    ok: true,
    service: "intermedia-exchange-connector",
    version: "0.5.0",
    features: ["search_elton_mail", "list_elton_attachments", "read_elton_attachment", "create_elton_email_draft", "finalize_elton_outreach_drafts"],
  });
});

app.all("/mcp", (_req, res) => res.status(404).json({ error: "Not found" }));
app.use(async (req, res, next) => {
  // Render-generated Base64 secrets can contain characters such as +, =, and /.
  // Compare the literal request path instead of compiling the secret as an
  // Express/path-to-regexp route pattern.
  if (req.path !== mcpPath) {
    next();
    return;
  }

  if (req.method !== "POST") {
    res.status(405).json({ error: "Method not allowed" });
    return;
  }

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

app.listen(config.port, (error) => {
  if (error) {
    console.error("Failed to start connector", error);
    process.exit(1);
  }
  console.log(`Intermedia Exchange connector listening on port ${config.port}`);
});
