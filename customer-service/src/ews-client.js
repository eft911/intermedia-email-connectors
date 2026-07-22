import { XMLParser } from "fast-xml-parser";
import { createHash } from "node:crypto";

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
  return responseMessages(body, operation)[0];
}

function responseMessages(body, operation) {
  const response = body?.[`${operation}Response`];
  const messages = response?.ResponseMessages;
  if (!messages) throw new Error(`EWS ${operation} response was missing ResponseMessages.`);
  const key = Object.keys(messages).find((name) => name.endsWith("ResponseMessage"));
  const found = key ? array(messages[key]) : [];
  if (!found.length) throw new Error(`EWS ${operation} response did not contain a response message.`);
  for (const message of found) {
    const responseCode = text(message.ResponseCode);
    if (message["@_ResponseClass"] !== "Success" || responseCode !== "NoError") {
      const detail = text(message.MessageText) || responseCode || "Unknown EWS error";
      throw new Error(`EWS ${operation} failed: ${detail}`);
    }
  }
  return found;
}

function normalizeMailbox(mailbox) {
  if (!mailbox) return null;
  return {
    name: text(mailbox.Name) || null,
    email: text(mailbox.EmailAddress) || null,
  };
}

function normalizeMailboxes(mailboxes) {
  return array(mailboxes?.Mailbox).map(normalizeMailbox).filter(Boolean);
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

function cleanBody(value) {
  return text(value)
    .replaceAll("\r", " ")
    .replaceAll("\n", " ")
    .replace(/\s+/g, " ")
    .trim();
}

function relevantExcerpt(body, query, maxLength = 320) {
  const cleaned = cleanBody(body);
  if (!cleaned) return null;
  const needle = query.toLocaleLowerCase();
  const index = cleaned.toLocaleLowerCase().indexOf(needle);
  const start = index < 0 ? 0 : Math.max(0, index - 100);
  const end = Math.min(cleaned.length, start + maxLength);
  return `${start > 0 ? "…" : ""}${cleaned.slice(start, end).trim()}${end < cleaned.length ? "…" : ""}`;
}

function queryFingerprint(query) {
  return createHash("sha256").update(query).digest("hex").slice(0, 16);
}

function encodeCursor(query, offsets) {
  return Buffer.from(JSON.stringify({ v: 1, q: queryFingerprint(query), offsets }), "utf8").toString("base64url");
}

function decodeCursor(cursor, query) {
  if (!cursor) return { inbox: 0, sentitems: 0, archive: 0 };
  try {
    const value = JSON.parse(Buffer.from(cursor, "base64url").toString("utf8"));
    const offsets = value?.offsets;
    if (value?.v !== 1 || value?.q !== queryFingerprint(query) || !offsets) throw new Error();
    for (const folder of ["inbox", "sentitems", "archive"]) {
      if (!Number.isSafeInteger(offsets[folder]) || offsets[folder] < 0) throw new Error();
    }
    return offsets;
  } catch {
    throw new Error("Invalid or expired search cursor for this query.");
  }
}

function safeAqsPhrase(query) {
  return query.replace(/["\\]/g, " ").replace(/\s+/g, " ").trim();
}

function itemDate(item) {
  return text(item?.DateTimeSent) || text(item?.DateTimeReceived) || text(item?.DateTimeCreated) || null;
}

function normalizeSearchDetails(item, summary, query) {
  return {
    date: itemDate(item) || summary.date,
    sender: normalizeMailbox(item?.From?.Mailbox || item?.Sender?.Mailbox),
    recipients: {
      to: normalizeMailboxes(item?.ToRecipients),
      cc: normalizeMailboxes(item?.CcRecipients),
      bcc: normalizeMailboxes(item?.BccRecipients),
    },
    subject: text(item?.Subject) || summary.subject || "(no subject)",
    folder: summary.folder,
    message_id: summary.id,
    excerpt: relevantExcerpt(item?.Body, query),
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
  constructor({ url, username, password, mailbox, fetchFn = fetch, timeoutMs = DEFAULT_TIMEOUT_MS, nowFn = () => new Date() }) {
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
    this.nowFn = nowFn;
    this.archiveFolderPromise = null;
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
    const pageSize = Math.min(Math.max(Number(maxResults) || 20, 1), 50);
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

  async resolveArchiveFolder() {
    if (!this.archiveFolderPromise) {
      this.archiveFolderPromise = this.request("FindFolder", `
<m:FindFolder Traversal="Deep">
  <m:FolderShape>
    <t:BaseShape>IdOnly</t:BaseShape>
    <t:AdditionalProperties><t:FieldURI FieldURI="folder:DisplayName"/></t:AdditionalProperties>
  </m:FolderShape>
  <m:Restriction>
    <t:IsEqualTo>
      <t:FieldURI FieldURI="folder:DisplayName"/>
      <t:FieldURIOrConstant><t:Constant Value="Archive"/></t:FieldURIOrConstant>
    </t:IsEqualTo>
  </m:Restriction>
  <m:ParentFolderIds>${folderIdXml(this.mailbox, "msgfolderroot")}</m:ParentFolderIds>
</m:FindFolder>`).then((body) => {
        const message = firstResponseMessage(body, "FindFolder");
        const folders = [
          ...array(message?.RootFolder?.Folders?.Folder),
          ...array(message?.RootFolder?.Folders?.SearchFolder),
        ];
        const archive = folders.find((folder) => text(folder?.DisplayName).toLocaleLowerCase() === "archive");
        const id = archive?.FolderId?.["@_Id"];
        if (!id) throw new Error("The Archive folder could not be found in the configured mailbox.");
        return { id, changeKey: archive?.FolderId?.["@_ChangeKey"] || null };
      }).catch((error) => {
        this.archiveFolderPromise = null;
        throw error;
      });
    }
    return this.archiveFolderPromise;
  }

  async searchFolder({ key, label, folderXml, query, cutoff, offset, pageSize, dateKeyword }) {
    const phrase = safeAqsPhrase(query);
    const date = cutoff.toISOString().slice(0, 10);
    const aqs = `(subject:"${phrase}" OR body:"${phrase}") AND ${dateKeyword}>=${date}`;
    const body = await this.request("FindItem", `
<m:FindItem Traversal="Shallow">
  <m:ItemShape>
    <t:BaseShape>IdOnly</t:BaseShape>
    <t:AdditionalProperties>
      <t:FieldURI FieldURI="item:Subject"/>
      <t:FieldURI FieldURI="item:DateTimeReceived"/>
      <t:FieldURI FieldURI="item:DateTimeSent"/>
      <t:FieldURI FieldURI="item:DateTimeCreated"/>
    </t:AdditionalProperties>
  </m:ItemShape>
  <m:IndexedPageItemView MaxEntriesReturned="${pageSize}" Offset="${offset}" BasePoint="Beginning"/>
  <m:QueryString>${xmlEscape(aqs)}</m:QueryString>
  <m:SortOrder><t:FieldOrder Order="Descending"><t:FieldURI FieldURI="item:DateTimeCreated"/></t:FieldOrder></m:SortOrder>
  <m:ParentFolderIds>${folderXml}</m:ParentFolderIds>
</m:FindItem>`);
    const message = firstResponseMessage(body, "FindItem");
    const root = message.RootFolder;
    const items = [
      ...array(root?.Items?.Message),
      ...array(root?.Items?.MeetingRequest),
    ].map((item) => ({
      id: item?.ItemId?.["@_Id"] || "",
      changeKey: item?.ItemId?.["@_ChangeKey"] || "",
      subject: text(item?.Subject) || "(no subject)",
      date: itemDate(item),
      folder: label,
      folderKey: key,
    })).filter((item) => item.id);
    return {
      items,
      hasMore: String(root?.["@_IncludesLastItemInRange"]).toLocaleLowerCase() !== "true",
    };
  }

  async getSearchDetails(summaries, query) {
    if (!summaries.length) return [];
    const ids = summaries.map((summary) => `<t:ItemId Id="${xmlEscape(summary.id)}"/>`).join("");
    const body = await this.request("GetItem", `
<m:GetItem>
  <m:ItemShape>
    <t:BaseShape>IdOnly</t:BaseShape>
    <t:BodyType>Text</t:BodyType>
    <t:AdditionalProperties>
      <t:FieldURI FieldURI="item:Subject"/>
      <t:FieldURI FieldURI="item:DateTimeReceived"/>
      <t:FieldURI FieldURI="item:DateTimeSent"/>
      <t:FieldURI FieldURI="item:DateTimeCreated"/>
      <t:FieldURI FieldURI="message:From"/>
      <t:FieldURI FieldURI="message:Sender"/>
      <t:FieldURI FieldURI="message:ToRecipients"/>
      <t:FieldURI FieldURI="message:CcRecipients"/>
      <t:FieldURI FieldURI="message:BccRecipients"/>
      <t:FieldURI FieldURI="item:Body"/>
    </t:AdditionalProperties>
  </m:ItemShape>
  <m:ItemIds>${ids}</m:ItemIds>
</m:GetItem>`);
    const messages = responseMessages(body, "GetItem");
    return summaries.map((summary, index) => {
      const response = messages[index];
      const item = response?.Items?.Message || response?.Items?.MeetingRequest;
      if (!item) throw new Error("EWS GetItem returned no email message for a search result.");
      return normalizeSearchDetails(item, summary, query);
    });
  }

  async searchMessages({ query, pageSize = 20, cursor = null }) {
    const normalizedQuery = String(query || "").trim();
    if (normalizedQuery.length < 2) throw new Error("Search query must contain at least 2 characters.");
    if (normalizedQuery.length > 200) throw new Error("Search query must not exceed 200 characters.");
    const size = Math.min(Math.max(Number(pageSize) || 20, 1), 50);
    const offsets = decodeCursor(cursor, normalizedQuery);
    const cutoff = new Date(this.nowFn().getTime() - 365 * 24 * 60 * 60 * 1000);
    const archive = await this.resolveArchiveFolder();
    const fetchSize = Math.min(size + 1, 50);
    const folders = [
      { key: "inbox", label: "Inbox", folderXml: folderIdXml(this.mailbox, "inbox"), dateKeyword: "received" },
      { key: "sentitems", label: "Sent Items", folderXml: folderIdXml(this.mailbox, "sentitems"), dateKeyword: "sent" },
      { key: "archive", label: "Archive", folderXml: `<t:FolderId Id="${xmlEscape(archive.id)}"/>`, dateKeyword: "received" },
    ];
    const pages = [];
    for (const folder of folders) {
      pages.push(await this.searchFolder({
        ...folder,
        query: normalizedQuery,
        cutoff,
        offset: offsets[folder.key],
        pageSize: fetchSize,
      }));
    }
    const merged = pages.flatMap((page) => page.items)
      .sort((a, b) => String(b.date || "").localeCompare(String(a.date || "")));
    const selected = merged.slice(0, size);
    const nextOffsets = { ...offsets };
    for (const item of selected) nextOffsets[item.folderKey] += 1;
    const hasMore = merged.length > selected.length || pages.some((page, index) =>
      page.hasMore || page.items.length > selected.filter((item) => item.folderKey === folders[index].key).length
    );
    const results = await this.getSearchDetails(selected, normalizedQuery);
    return {
      mailbox: this.mailbox,
      query: normalizedQuery,
      searched_since: cutoff.toISOString(),
      folders: folders.map((folder) => folder.label),
      page_size: size,
      results,
      has_more: hasMore,
      next_cursor: hasMore ? encodeCursor(normalizedQuery, nextOffsets) : null,
    };
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

export const _test = { parseSoap, firstResponseMessage, responseMessages, normalizeMessage, xmlEscape, relevantExcerpt, encodeCursor, decodeCursor };
