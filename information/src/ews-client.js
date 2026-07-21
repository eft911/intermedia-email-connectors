import { XMLParser } from "fast-xml-parser";

const SOAP_NS = "http://schemas.xmlsoap.org/soap/envelope/";
const MESSAGES_NS = "http://schemas.microsoft.com/exchange/services/2006/messages";
const TYPES_NS = "http://schemas.microsoft.com/exchange/services/2006/types";
const DEFAULT_TIMEOUT_MS = 25_000;
export const PROCESSED_CATEGORY = "GPT Drafted";

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  removeNSPrefix: true,
  trimValues: true,
  parseTagValue: false,
});

function xmlEscape(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;");
}

function array(value) {
  if (value == null) return [];
  return Array.isArray(value) ? value : [value];
}

function text(value) {
  if (value == null) return "";
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
    return String(value);
  }
  if (typeof value === "object" && "#text" in value) return String(value["#text"]);
  return "";
}

function firstResponseMessage(body, operation) {
  const response = body?.[`${operation}Response`];
  const messages = response?.ResponseMessages;
  if (!messages) throw new Error(`EWS ${operation} response was missing ResponseMessages.`);
  const key = Object.keys(messages).find((name) => name.endsWith("ResponseMessage"));
  const message = key ? array(messages[key])[0] : null;
  if (!message) throw new Error(`EWS ${operation} response did not contain a response message.`);
  const responseCode = text(message.ResponseCode);
  if (message["@_ResponseClass"] !== "Success" || responseCode !== "NoError") {
    const detail = text(message.MessageText) || responseCode || "Unknown EWS error";
    throw new Error(`EWS ${operation} failed: ${detail}`);
  }
  return message;
}

function normalizeMailbox(mailbox) {
  if (!mailbox) return null;
  return {
    name: text(mailbox.Name) || null,
    email: text(mailbox.EmailAddress) || null,
  };
}

function normalizeCategories(categories) {
  return array(categories?.String).map(text).filter(Boolean);
}

function normalizeMessage(item) {
  const itemId = item?.ItemId;
  return {
    id: itemId?.["@_Id"] || "",
    change_key: itemId?.["@_ChangeKey"] || "",
    subject: text(item?.Subject) || "(no subject)",
    from: normalizeMailbox(item?.From?.Mailbox),
    received_at: text(item?.DateTimeReceived) || null,
    is_read: text(item?.IsRead).toLowerCase() === "true",
    has_attachments: text(item?.HasAttachments).toLowerCase() === "true",
    categories: normalizeCategories(item?.Categories),
    body_type: item?.Body?.["@_BodyType"] || null,
    body: text(item?.Body) || null,
  };
}

function parseSoap(xml) {
  const parsed = parser.parse(xml);
  const envelope = parsed?.Envelope;
  const body = envelope?.Body;
  const fault = body?.Fault;
  if (fault) {
    const detail = text(fault.faultstring) || text(fault.Reason?.Text) || "Unknown SOAP fault";
    throw new Error(`EWS SOAP fault: ${detail}`);
  }
  if (!body) throw new Error("EWS returned an invalid SOAP response.");
  return body;
}

function folderIdXml(mailbox, name) {
  return `<t:DistinguishedFolderId Id="${xmlEscape(name)}"><t:Mailbox><t:EmailAddress>${xmlEscape(mailbox)}</t:EmailAddress></t:Mailbox></t:DistinguishedFolderId>`;
}

export class EwsClient {
  constructor({ url, username, password, mailbox, fetchFn = fetch, timeoutMs = DEFAULT_TIMEOUT_MS }) {
    this.url = new URL(url);
    if (this.url.protocol !== "https:" || !this.url.hostname.toLowerCase().endsWith(".serverdata.net")) {
      throw new Error("EWS_URL must be an HTTPS serverdata.net EWS endpoint.");
    }
    if (!this.url.pathname.toLowerCase().endsWith("/ews/exchange.asmx")) {
      throw new Error("EWS_URL must end in /EWS/Exchange.asmx.");
    }
    if (!username || !password || !mailbox) throw new Error("EWS credentials and SHARED_MAILBOX are required.");
    this.username = username;
    this.password = password;
    this.mailbox = mailbox;
    this.fetchFn = fetchFn;
    this.timeoutMs = timeoutMs;
  }

  envelope(operationXml) {
    return `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="${SOAP_NS}" xmlns:m="${MESSAGES_NS}" xmlns:t="${TYPES_NS}">
  <soap:Header><t:RequestServerVersion Version="Exchange2013_SP1"/></soap:Header>
  <soap:Body>${operationXml}</soap:Body>
</soap:Envelope>`;
  }

  async request(operation, operationXml) {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), this.timeoutMs);
    let response;
    try {
      response = await this.fetchFn(this.url, {
        method: "POST",
        headers: {
          Accept: "text/xml",
          "Content-Type": "text/xml; charset=utf-8",
          Authorization: `Basic ${Buffer.from(`${this.username}:${this.password}`).toString("base64")}`,
          "User-Agent": "MeTooShoes-Intermedia-MCP/0.1",
        },
        body: this.envelope(operationXml),
        signal: controller.signal,
      });
    } catch (error) {
      if (error?.name === "AbortError") throw new Error(`EWS ${operation} timed out.`);
      throw new Error(`EWS ${operation} request failed.`);
    } finally {
      clearTimeout(timer);
    }

    const xml = await response.text();
    if (!response.ok) {
      if (response.status === 401) throw new Error("EWS authentication failed. Check the username, password, and Exchange 2FA policy.");
      throw new Error(`EWS ${operation} returned HTTP ${response.status}.`);
    }
    return parseSoap(xml);
  }

  async healthCheck() {
    const body = await this.request("GetFolder", `
<m:GetFolder>
  <m:FolderShape><t:BaseShape>Default</t:BaseShape></m:FolderShape>
  <m:FolderIds>${folderIdXml(this.mailbox, "inbox")}</m:FolderIds>
</m:GetFolder>`);
    const message = firstResponseMessage(body, "GetFolder");
    const folder = message?.Folders?.Folder || message?.Folders?.CalendarFolder;
    return {
      ok: true,
      mailbox: this.mailbox,
      total_count: Number(text(folder?.TotalCount) || 0),
      unread_count: Number(text(folder?.UnreadCount) || 0),
    };
  }

  async listMessages({ unreadOnly = true, sinceHours = 72, maxResults = 20, includeProcessed = false } = {}) {
    const conditions = [];
    if (unreadOnly) {
      conditions.push(`<t:IsEqualTo><t:FieldURI FieldURI="message:IsRead"/><t:FieldURIOrConstant><t:Constant Value="false"/></t:FieldURIOrConstant></t:IsEqualTo>`);
    }
    if (sinceHours > 0) {
      const since = new Date(Date.now() - sinceHours * 60 * 60 * 1000).toISOString();
      conditions.push(`<t:IsGreaterThanOrEqualTo><t:FieldURI FieldURI="item:DateTimeReceived"/><t:FieldURIOrConstant><t:Constant Value="${since}"/></t:FieldURIOrConstant></t:IsGreaterThanOrEqualTo>`);
    }
    const restriction = conditions.length
      ? `<m:Restriction>${conditions.length === 1 ? conditions[0] : `<t:And>${conditions.join("")}</t:And>`}</m:Restriction>`
      : "";
    const pageSize = Math.min(Math.max(Number(maxResults) || 20, 1), 500);
    const body = await this.request("FindItem", `
<m:FindItem Traversal="Shallow">
  <m:ItemShape>
    <t:BaseShape>IdOnly</t:BaseShape>
    <t:AdditionalProperties>
      <t:FieldURI FieldURI="item:Subject"/>
      <t:FieldURI FieldURI="message:From"/>
      <t:FieldURI FieldURI="item:DateTimeReceived"/>
      <t:FieldURI FieldURI="message:IsRead"/>
      <t:FieldURI FieldURI="item:HasAttachments"/>
      <t:FieldURI FieldURI="item:Categories"/>
    </t:AdditionalProperties>
  </m:ItemShape>
  <m:IndexedPageItemView MaxEntriesReturned="${pageSize}" Offset="0" BasePoint="Beginning"/>
  ${restriction}
  <m:SortOrder><t:FieldOrder Order="Descending"><t:FieldURI FieldURI="item:DateTimeReceived"/></t:FieldOrder></m:SortOrder>
  <m:ParentFolderIds>${folderIdXml(this.mailbox, "inbox")}</m:ParentFolderIds>
</m:FindItem>`);
    const responseMessage = firstResponseMessage(body, "FindItem");
    const root = responseMessage.RootFolder;
    const items = [
      ...array(root?.Items?.Message),
      ...array(root?.Items?.MeetingRequest),
    ].map(normalizeMessage).filter((message) => message.id);
    const filtered = includeProcessed
      ? items
      : items.filter((message) => !message.categories.includes(PROCESSED_CATEGORY));
    return {
      mailbox: this.mailbox,
      total_items_in_view: Number(root?.["@_TotalItemsInView"] || filtered.length),
      messages: filtered,
    };
  }

  async getMessage(messageId) {
    const body = await this.request("GetItem", `
<m:GetItem>
  <m:ItemShape>
    <t:BaseShape>IdOnly</t:BaseShape>
    <t:BodyType>Text</t:BodyType>
    <t:AdditionalProperties>
      <t:FieldURI FieldURI="item:Subject"/>
      <t:FieldURI FieldURI="message:From"/>
      <t:FieldURI FieldURI="item:DateTimeReceived"/>
      <t:FieldURI FieldURI="message:IsRead"/>
      <t:FieldURI FieldURI="item:HasAttachments"/>
      <t:FieldURI FieldURI="item:Categories"/>
      <t:FieldURI FieldURI="item:Body"/>
    </t:AdditionalProperties>
  </m:ItemShape>
  <m:ItemIds><t:ItemId Id="${xmlEscape(messageId)}"/></m:ItemIds>
</m:GetItem>`);
    const responseMessage = firstResponseMessage(body, "GetItem");
    const item = responseMessage?.Items?.Message || responseMessage?.Items?.MeetingRequest;
    if (!item) throw new Error("EWS GetItem returned no email message.");
    return normalizeMessage(item);
  }

  async createReplyDraft({ messageId, textContent }) {
    const source = await this.getMessage(messageId);
    if (source.categories.includes(PROCESSED_CATEGORY)) {
      return { created: false, duplicate_prevented: true, source_message: source };
    }
    const body = await this.request("CreateItem", `
<m:CreateItem MessageDisposition="SaveOnly">
  <m:SavedItemFolderId>${folderIdXml(this.mailbox, "drafts")}</m:SavedItemFolderId>
  <m:Items>
    <t:ReplyToItem>
      <t:ReferenceItemId Id="${xmlEscape(source.id)}" ChangeKey="${xmlEscape(source.change_key)}"/>
      <t:NewBodyContent BodyType="Text">${xmlEscape(textContent)}</t:NewBodyContent>
    </t:ReplyToItem>
  </m:Items>
</m:CreateItem>`);
    const responseMessage = firstResponseMessage(body, "CreateItem");
    const draftItemId = responseMessage?.Items?.Message?.ItemId;
    let categoryAdded = false;
    let warning = null;
    try {
      await this.addProcessedCategory(source);
      categoryAdded = true;
    } catch {
      warning = `The draft was saved, but the ${PROCESSED_CATEGORY} category could not be added. Review the source message before retrying to avoid a duplicate draft.`;
    }
    return {
      created: true,
      duplicate_prevented: false,
      mailbox: this.mailbox,
      draft_id: draftItemId?.["@_Id"] || null,
      source_message_id: source.id,
      source_subject: source.subject,
      category_added: categoryAdded ? PROCESSED_CATEGORY : null,
      warning,
    };
  }

  async addProcessedCategory(source) {
    const categories = [...new Set([...source.categories, PROCESSED_CATEGORY])];
    const categoryXml = categories.map((category) => `<t:String>${xmlEscape(category)}</t:String>`).join("");
    const body = await this.request("UpdateItem", `
<m:UpdateItem ConflictResolution="AutoResolve" MessageDisposition="SaveOnly">
  <m:ItemChanges>
    <t:ItemChange>
      <t:ItemId Id="${xmlEscape(source.id)}" ChangeKey="${xmlEscape(source.change_key)}"/>
      <t:Updates>
        <t:SetItemField>
          <t:FieldURI FieldURI="item:Categories"/>
          <t:Message><t:Categories>${categoryXml}</t:Categories></t:Message>
        </t:SetItemField>
      </t:Updates>
    </t:ItemChange>
  </m:ItemChanges>
</m:UpdateItem>`);
    firstResponseMessage(body, "UpdateItem");
  }
}

export function loadConfig(env = process.env) {
  const required = ["EWS_URL", "EWS_USERNAME", "EWS_PASSWORD", "SHARED_MAILBOX", "MCP_PATH_SECRET"];
  const missing = required.filter((key) => !env[key]);
  if (missing.length) throw new Error(`Missing required environment variables: ${missing.join(", ")}`);
  if (env.MCP_PATH_SECRET.length < 24) throw new Error("MCP_PATH_SECRET must be at least 24 characters.");
  return {
    ews: {
      url: env.EWS_URL,
      username: env.EWS_USERNAME,
      password: env.EWS_PASSWORD,
      mailbox: env.SHARED_MAILBOX,
    },
    mcpPathSecret: env.MCP_PATH_SECRET,
    port: Number(env.PORT || 3000),
  };
}

export const _test = { parseSoap, firstResponseMessage, normalizeMessage, xmlEscape };
