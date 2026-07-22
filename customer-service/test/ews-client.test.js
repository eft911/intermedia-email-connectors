import test from "node:test";
import assert from "node:assert/strict";
import { EwsClient, PROCESSED_CATEGORY, _test } from "../src/ews-client.js";

function soap(inner) {
  return `<?xml version="1.0"?><s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"><s:Body>${inner}</s:Body></s:Envelope>`;
}

function response(xml, status = 200) {
  return new Response(xml, { status, headers: { "content-type": "text/xml" } });
}

function clientWith(responses, requests = []) {
  return new EwsClient({
    url: "https://east.exch028.serverdata.net/EWS/Exchange.asmx",
    username: "service@example.com",
    password: "secret",
    mailbox: "customerservice@example.com",
    fetchFn: async (_url, options) => {
      requests.push(options.body);
      const next = responses.shift();
      if (!next) throw new Error("Unexpected request");
      return next;
    },
  });
}

test("healthCheck parses shared inbox counts", async () => {
  const client = clientWith([response(soap(`
    <m:GetFolderResponse><m:ResponseMessages><m:GetFolderResponseMessage ResponseClass="Success"><m:ResponseCode>NoError</m:ResponseCode><m:Folders><t:Folder><t:TotalCount>42</t:TotalCount><t:UnreadCount>7</t:UnreadCount></t:Folder></m:Folders></m:GetFolderResponseMessage></m:ResponseMessages></m:GetFolderResponse>`))]);
  assert.deepEqual(await client.healthCheck(), {
    ok: true,
    mailbox: "customerservice@example.com",
    total_count: 42,
    unread_count: 7,
  });
});

test("listMessages excludes messages already categorized as drafted", async () => {
  const client = clientWith([response(soap(`
    <m:FindItemResponse><m:ResponseMessages><m:FindItemResponseMessage ResponseClass="Success"><m:ResponseCode>NoError</m:ResponseCode><m:RootFolder TotalItemsInView="2"><t:Items>
      <t:Message><t:ItemId Id="one" ChangeKey="ck1"/><t:Subject>Order one</t:Subject><t:From><t:Mailbox><t:Name>Jane</t:Name><t:EmailAddress>jane@example.com</t:EmailAddress></t:Mailbox></t:From><t:DateTimeReceived>2026-07-21T12:00:00Z</t:DateTimeReceived><t:IsRead>false</t:IsRead><t:HasAttachments>false</t:HasAttachments></t:Message>
      <t:Message><t:ItemId Id="two" ChangeKey="ck2"/><t:Subject>Order two</t:Subject><t:Categories><t:String>${PROCESSED_CATEGORY}</t:String></t:Categories><t:IsRead>false</t:IsRead><t:HasAttachments>false</t:HasAttachments></t:Message>
    </t:Items></m:RootFolder></m:FindItemResponseMessage></m:ResponseMessages></m:FindItemResponse>`))]);
  const result = await client.listMessages();
  assert.equal(result.messages.length, 1);
  assert.equal(result.messages[0].id, "one");
});

test("createReplyDraft saves only and categorizes source", async () => {
  const requests = [];
  const client = clientWith([
    response(soap(`<m:GetItemResponse><m:ResponseMessages><m:GetItemResponseMessage ResponseClass="Success"><m:ResponseCode>NoError</m:ResponseCode><m:Items><t:Message><t:ItemId Id="source-id" ChangeKey="source-ck"/><t:Subject>Order 123</t:Subject><t:Categories><t:String>Customer</t:String></t:Categories><t:IsRead>false</t:IsRead><t:Body BodyType="Text">Help</t:Body></t:Message></m:Items></m:GetItemResponseMessage></m:ResponseMessages></m:GetItemResponse>`)),
    response(soap(`<m:CreateItemResponse><m:ResponseMessages><m:CreateItemResponseMessage ResponseClass="Success"><m:ResponseCode>NoError</m:ResponseCode><m:Items><t:Message><t:ItemId Id="draft-id" ChangeKey="draft-ck"/></t:Message></m:Items></m:CreateItemResponseMessage></m:ResponseMessages></m:CreateItemResponse>`)),
    response(soap(`<m:UpdateItemResponse><m:ResponseMessages><m:UpdateItemResponseMessage ResponseClass="Success"><m:ResponseCode>NoError</m:ResponseCode></m:UpdateItemResponseMessage></m:ResponseMessages></m:UpdateItemResponse>`)),
  ], requests);
  const result = await client.createReplyDraft({ messageId: "source-id", textContent: "We are happy to help." });
  assert.equal(result.created, true);
  assert.equal(result.draft_id, "draft-id");
  assert.equal(result.category_added, PROCESSED_CATEGORY);
  assert.match(requests[1], /MessageDisposition="SaveOnly"/);
  assert.match(requests[1], /We are happy to help\./);
  assert.match(requests[2], /GPT Drafted/);
});

test("searchMessages searches Inbox, Sent Items, and Archive with pagination", async () => {
  const requests = [];
  const client = new EwsClient({
    url: "https://east.exch028.serverdata.net/EWS/Exchange.asmx",
    username: "service@example.com",
    password: "secret",
    mailbox: "etucker@metooshoes.com",
    nowFn: () => new Date("2026-07-22T12:00:00Z"),
    fetchFn: async (_url, options) => {
      requests.push(options.body);
      const next = responses.shift();
      if (!next) throw new Error("Unexpected request");
      return next;
    },
  });
  const responses = [
    response(soap(`<m:FindFolderResponse><m:ResponseMessages><m:FindFolderResponseMessage ResponseClass="Success"><m:ResponseCode>NoError</m:ResponseCode><m:RootFolder IncludesLastItemInRange="true" TotalItemsInView="1"><t:Folders><t:Folder><t:FolderId Id="archive-folder-id" ChangeKey="archive-ck"/><t:DisplayName>Archive</t:DisplayName></t:Folder></t:Folders></m:RootFolder></m:FindFolderResponseMessage></m:ResponseMessages></m:FindFolderResponse>`)),
    response(soap(`<m:FindItemResponse><m:ResponseMessages><m:FindItemResponseMessage ResponseClass="Success"><m:ResponseCode>NoError</m:ResponseCode><m:RootFolder IncludesLastItemInRange="true" TotalItemsInView="1"><t:Items><t:Message><t:ItemId Id="inbox-id" ChangeKey="inbox-ck"/><t:Subject>Launch plan inbox</t:Subject><t:DateTimeReceived>2026-07-21T10:00:00Z</t:DateTimeReceived><t:DateTimeCreated>2026-07-21T10:00:00Z</t:DateTimeCreated></t:Message></t:Items></m:RootFolder></m:FindItemResponseMessage></m:ResponseMessages></m:FindItemResponse>`)),
    response(soap(`<m:FindItemResponse><m:ResponseMessages><m:FindItemResponseMessage ResponseClass="Success"><m:ResponseCode>NoError</m:ResponseCode><m:RootFolder IncludesLastItemInRange="true" TotalItemsInView="1"><t:Items><t:Message><t:ItemId Id="sent-id" ChangeKey="sent-ck"/><t:Subject>Re: Launch plan</t:Subject><t:DateTimeSent>2026-07-22T09:00:00Z</t:DateTimeSent><t:DateTimeCreated>2026-07-22T09:00:00Z</t:DateTimeCreated></t:Message></t:Items></m:RootFolder></m:FindItemResponseMessage></m:ResponseMessages></m:FindItemResponse>`)),
    response(soap(`<m:FindItemResponse><m:ResponseMessages><m:FindItemResponseMessage ResponseClass="Success"><m:ResponseCode>NoError</m:ResponseCode><m:RootFolder IncludesLastItemInRange="true" TotalItemsInView="1"><t:Items><t:Message><t:ItemId Id="archive-id" ChangeKey="archive-ck"/><t:Subject>Older launch plan</t:Subject><t:DateTimeReceived>2026-07-20T08:00:00Z</t:DateTimeReceived><t:DateTimeCreated>2026-07-20T08:00:00Z</t:DateTimeCreated></t:Message></t:Items></m:RootFolder></m:FindItemResponseMessage></m:ResponseMessages></m:FindItemResponse>`)),
    response(soap(`<m:GetItemResponse><m:ResponseMessages>
      <m:GetItemResponseMessage ResponseClass="Success"><m:ResponseCode>NoError</m:ResponseCode><m:Items><t:Message><t:ItemId Id="sent-id"/><t:Subject>Re: Launch plan</t:Subject><t:DateTimeSent>2026-07-22T09:00:00Z</t:DateTimeSent><t:From><t:Mailbox><t:Name>Elton</t:Name><t:EmailAddress>etucker@metooshoes.com</t:EmailAddress></t:Mailbox></t:From><t:ToRecipients><t:Mailbox><t:Name>Sam</t:Name><t:EmailAddress>sam@example.com</t:EmailAddress></t:Mailbox></t:ToRecipients><t:CcRecipients><t:Mailbox><t:EmailAddress>team@example.com</t:EmailAddress></t:Mailbox></t:CcRecipients><t:Body BodyType="Text">Here is the final launch plan for next week.</t:Body></t:Message></m:Items></m:GetItemResponseMessage>
      <m:GetItemResponseMessage ResponseClass="Success"><m:ResponseCode>NoError</m:ResponseCode><m:Items><t:Message><t:ItemId Id="inbox-id"/><t:Subject>Launch plan inbox</t:Subject><t:DateTimeReceived>2026-07-21T10:00:00Z</t:DateTimeReceived><t:From><t:Mailbox><t:Name>Jane</t:Name><t:EmailAddress>jane@example.com</t:EmailAddress></t:Mailbox></t:From><t:ToRecipients><t:Mailbox><t:EmailAddress>etucker@metooshoes.com</t:EmailAddress></t:Mailbox></t:ToRecipients><t:Body BodyType="Text">Can you review the launch plan?</t:Body></t:Message></m:Items></m:GetItemResponseMessage>
    </m:ResponseMessages></m:GetItemResponse>`)),
  ];

  const result = await client.searchMessages({ query: "launch plan", pageSize: 2 });
  assert.equal(result.mailbox, "etucker@metooshoes.com");
  assert.equal(result.searched_since, "2025-07-22T12:00:00.000Z");
  assert.deepEqual(result.folders, ["Inbox", "Sent Items", "Archive"]);
  assert.equal(result.results.length, 2);
  assert.equal(result.results[0].folder, "Sent Items");
  assert.equal(result.results[0].message_id, "sent-id");
  assert.equal(result.results[0].sender.email, "etucker@metooshoes.com");
  assert.equal(result.results[0].recipients.to[0].email, "sam@example.com");
  assert.equal(result.results[0].recipients.cc[0].email, "team@example.com");
  assert.match(result.results[0].excerpt, /launch plan/);
  assert.equal(result.has_more, true);
  assert.ok(result.next_cursor);
  assert.deepEqual(_test.decodeCursor(result.next_cursor, "launch plan"), {
    inbox: 1,
    sentitems: 1,
    archive: 0,
  });
  assert.match(requests[0], /Traversal="Deep"/);
  assert.match(requests[0], /DisplayName/);
  assert.match(requests[1], /subject:&quot;launch plan&quot; OR body:&quot;launch plan&quot;/);
  assert.match(requests[1], /received:07\/22\/2025\.\.07\/22\/2026/);
  assert.match(requests[2], /sent:07\/22\/2025\.\.07\/22\/2026/);
  assert.match(requests[3], /FolderId Id="archive-folder-id"/);
  assert.match(requests[4], /BodyType>Text/);
  assert.equal(requests.some((request) => /CreateItem|UpdateItem|DeleteItem|MoveItem/.test(request)), false);
});

test("search cursors are bound to the original query", () => {
  const cursor = _test.encodeCursor("first query", { inbox: 1, sentitems: 2, archive: 3 });
  assert.throws(() => _test.decodeCursor(cursor, "different query"), /Invalid or expired/);
});

test("rejects non-Intermedia EWS URLs", () => {
  assert.throws(() => new EwsClient({
    url: "https://evil.example/EWS/Exchange.asmx",
    username: "u",
    password: "p",
    mailbox: "m@example.com",
  }), /serverdata\.net/);
});
