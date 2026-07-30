import test from "node:test";
import assert from "node:assert/strict";
import { EwsClient, PROCESSED_CATEGORY, SUPPRESSED_RECIPIENTS, SUPPRESSED_RECIPIENT_DOMAINS, _test } from "../src/ews-client.js";

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

test("createEmailDraft saves a new outbound message without sending", async () => {
  const requests = [];
  const client = clientWith([
    response(soap(`<m:CreateItemResponse><m:ResponseMessages><m:CreateItemResponseMessage ResponseClass="Success"><m:ResponseCode>NoError</m:ResponseCode><m:Items><t:Message><t:ItemId Id="new-draft-id" ChangeKey="draft-ck"/></t:Message></m:Items></m:CreateItemResponseMessage></m:ResponseMessages></m:CreateItemResponse>`)),
  ], requests);
  const result = await client.createEmailDraft({
    to: ["buyer@example.com"],
    cc: ["sales@example.com"],
    bcc: [],
    subject: "A & B footwear",
    textContent: "Hello <Buyer>,\n\nA fresh idea for you.",
  });
  assert.equal(result.created, true);
  assert.equal(result.draft_id, "new-draft-id");
  assert.deepEqual(result.to, ["buyer@example.com"]);
  assert.match(requests[0], /MessageDisposition="SaveOnly"/);
  assert.match(requests[0], /DistinguishedFolderId Id="drafts"/);
  assert.match(requests[0], /buyer@example\.com/);
  assert.match(requests[0], /sales@example\.com/);
  assert.match(requests[0], /A &amp; B footwear/);
  assert.match(requests[0], /Hello &lt;Buyer&gt;/);
  assert.doesNotMatch(requests[0], /SendOnly|SendAndSaveCopy/);
});

test("suppressed recipients cannot be drafted", async () => {
  const client = clientWith([]);
  assert.equal(SUPPRESSED_RECIPIENTS.has("spinto@chineselaundry.com"), true);
  await assert.rejects(() => client.createEmailDraft({
    to: ["spinto@chineselaundry.com"],
    subject: "Outreach",
    textContent: "Hello",
  }), /Suppressed recipient/);
  assert.equal(SUPPRESSED_RECIPIENTS.has("shkatan@aol.com"), true);
  await assert.rejects(() => client.createEmailDraft({
    to: ["shkatan@aol.com"],
    subject: "Outreach",
    textContent: "Hello",
  }), /Suppressed recipient/);
  for (const email of [
    "jbancroft@propetusa.com",
    "jbrookings@propetusa.com",
    "todd.combs@keenfootwear.com",
    "lbalfour@superfeet.com",
  ]) {
    assert.equal(SUPPRESSED_RECIPIENTS.has(email), true);
  }
  assert.equal(SUPPRESSED_RECIPIENT_DOMAINS.has("pb5star.com"), true);
  await assert.rejects(() => client.createEmailDraft({
    to: ["anyone@PB5STAR.com"],
    subject: "Outreach",
    textContent: "Hello",
  }), /Suppressed recipient/);
});

test("finalizeOutreachDrafts validates, updates signatures, and sends reviewed drafts", async () => {
  const requests = [];
  const client = clientWith([
    response(soap(`<m:FindItemResponse><m:ResponseMessages><m:FindItemResponseMessage ResponseClass="Success"><m:ResponseCode>NoError</m:ResponseCode><m:RootFolder IncludesLastItemInRange="true" TotalItemsInView="2"><t:Items>
      <t:Message><t:ItemId Id="draft-1" ChangeKey="ck-1"/><t:Subject>Footwear Development &amp; Production</t:Subject></t:Message>
      <t:Message><t:ItemId Id="draft-2" ChangeKey="ck-2"/><t:Subject>Footwear Development &amp; Production</t:Subject></t:Message>
    </t:Items></m:RootFolder></m:FindItemResponseMessage></m:ResponseMessages></m:FindItemResponse>`)),
    response(soap(`<m:GetItemResponse><m:ResponseMessages>
      <m:GetItemResponseMessage ResponseClass="Success"><m:ResponseCode>NoError</m:ResponseCode><m:Items><t:Message><t:ItemId Id="draft-1" ChangeKey="ck-1"/><t:Subject>Footwear Development &amp; Production</t:Subject><t:Body BodyType="Text">Hi A

Thanks,
Elton</t:Body><t:ToRecipients><t:Mailbox><t:EmailAddress>a@example.com</t:EmailAddress></t:Mailbox></t:ToRecipients></t:Message></m:Items></m:GetItemResponseMessage>
      <m:GetItemResponseMessage ResponseClass="Success"><m:ResponseCode>NoError</m:ResponseCode><m:Items><t:Message><t:ItemId Id="draft-2" ChangeKey="ck-2"/><t:Subject>Footwear Development &amp; Production</t:Subject><t:Body BodyType="Text">Hi B

Thanks,
Elton</t:Body><t:ToRecipients><t:Mailbox><t:EmailAddress>b@example.com</t:EmailAddress></t:Mailbox></t:ToRecipients></t:Message></m:Items></m:GetItemResponseMessage>
    </m:ResponseMessages></m:GetItemResponse>`)),
    response(soap(`<m:UpdateItemResponse><m:ResponseMessages><m:UpdateItemResponseMessage ResponseClass="Success"><m:ResponseCode>NoError</m:ResponseCode><m:Items><t:Message><t:ItemId Id="draft-1" ChangeKey="new-ck-1"/></t:Message></m:Items></m:UpdateItemResponseMessage></m:ResponseMessages></m:UpdateItemResponse>`)),
    response(soap(`<m:UpdateItemResponse><m:ResponseMessages><m:UpdateItemResponseMessage ResponseClass="Success"><m:ResponseCode>NoError</m:ResponseCode><m:Items><t:Message><t:ItemId Id="draft-2" ChangeKey="new-ck-2"/></t:Message></m:Items></m:UpdateItemResponseMessage></m:ResponseMessages></m:UpdateItemResponse>`)),
    response(soap(`<m:SendItemResponse><m:ResponseMessages><m:SendItemResponseMessage ResponseClass="Success"><m:ResponseCode>NoError</m:ResponseCode></m:SendItemResponseMessage></m:ResponseMessages></m:SendItemResponse>`)),
    response(soap(`<m:SendItemResponse><m:ResponseMessages><m:SendItemResponseMessage ResponseClass="Success"><m:ResponseCode>NoError</m:ResponseCode></m:SendItemResponseMessage></m:ResponseMessages></m:SendItemResponse>`)),
  ], requests);
  const result = await client.finalizeOutreachDrafts({
    subject: "Footwear Development & Production",
    expectedCount: 2,
    oldSignature: "\nElton",
    newSignature: "\nElton Tucker",
  });
  assert.equal(result.sent_count, 2);
  assert.equal(result.failed_count, 0);
  assert.match(requests[0], /DistinguishedFolderId Id="drafts"/);
  assert.match(requests[0], /Footwear Development &amp; Production/);
  assert.match(requests[2], /Elton Tucker/);
  assert.match(requests[4], /SaveItemToFolder="true"/);
  assert.match(requests[4], /DistinguishedFolderId Id="sentitems"/);
});

test("finalizeOutreachDrafts refuses a count mismatch before any write", async () => {
  const requests = [];
  const client = clientWith([
    response(soap(`<m:FindItemResponse><m:ResponseMessages><m:FindItemResponseMessage ResponseClass="Success"><m:ResponseCode>NoError</m:ResponseCode><m:RootFolder IncludesLastItemInRange="true" TotalItemsInView="1"><t:Items><t:Message><t:ItemId Id="draft-1" ChangeKey="ck-1"/><t:Subject>Footwear Development</t:Subject></t:Message></t:Items></m:RootFolder></m:FindItemResponseMessage></m:ResponseMessages></m:FindItemResponse>`)),
  ], requests);
  await assert.rejects(() => client.finalizeOutreachDrafts({
    subject: "Footwear Development",
    expectedCount: 2,
    oldSignature: "\nElton",
    newSignature: "\nElton Tucker",
  }), /Expected 2 matching drafts but found 1/);
  assert.equal(requests.length, 1);
  assert.equal(requests.some((request) => /UpdateItem|SendItem/.test(request)), false);
});

test("finalizeOutreachDrafts skips an invalid draft and sends valid drafts", async () => {
  const requests = [];
  const client = clientWith([
    response(soap(`<m:FindItemResponse><m:ResponseMessages><m:FindItemResponseMessage ResponseClass="Success"><m:ResponseCode>NoError</m:ResponseCode><m:RootFolder IncludesLastItemInRange="true" TotalItemsInView="2"><t:Items>
      <t:Message><t:ItemId Id="valid" ChangeKey="ck-valid"/><t:Subject>Outreach</t:Subject></t:Message>
      <t:Message><t:ItemId Id="invalid" ChangeKey="ck-invalid"/><t:Subject>Outreach</t:Subject></t:Message>
    </t:Items></m:RootFolder></m:FindItemResponseMessage></m:ResponseMessages></m:FindItemResponse>`)),
    response(soap(`<m:GetItemResponse><m:ResponseMessages>
      <m:GetItemResponseMessage ResponseClass="Success"><m:ResponseCode>NoError</m:ResponseCode><m:Items><t:Message><t:ItemId Id="valid" ChangeKey="ck-valid"/><t:Body BodyType="Text">Thanks,
Elton</t:Body><t:ToRecipients><t:Mailbox><t:EmailAddress>valid@example.com</t:EmailAddress></t:Mailbox></t:ToRecipients></t:Message></m:Items></m:GetItemResponseMessage>
      <m:GetItemResponseMessage ResponseClass="Success"><m:ResponseCode>NoError</m:ResponseCode><m:Items><t:Message><t:ItemId Id="invalid" ChangeKey="ck-invalid"/><t:Body BodyType="Text">Changed signature</t:Body></t:Message></m:Items></m:GetItemResponseMessage>
    </m:ResponseMessages></m:GetItemResponse>`)),
    response(soap(`<m:UpdateItemResponse><m:ResponseMessages><m:UpdateItemResponseMessage ResponseClass="Success"><m:ResponseCode>NoError</m:ResponseCode><m:Items><t:Message><t:ItemId Id="valid" ChangeKey="new-ck"/></t:Message></m:Items></m:UpdateItemResponseMessage></m:ResponseMessages></m:UpdateItemResponse>`)),
    response(soap(`<m:SendItemResponse><m:ResponseMessages><m:SendItemResponseMessage ResponseClass="Success"><m:ResponseCode>NoError</m:ResponseCode></m:SendItemResponseMessage></m:ResponseMessages></m:SendItemResponse>`)),
  ], requests);
  const result = await client.finalizeOutreachDrafts({
    subject: "Outreach",
    expectedCount: 2,
    oldSignature: "Elton",
    newSignature: "Elton Tucker",
  });
  assert.equal(result.sent_count, 1);
  assert.equal(result.failed_count, 1);
  assert.equal(result.failures[0].recipient, null);
  assert.match(result.failures[0].error, /signature.*recipient/);
  assert.equal(requests.filter((request) => /UpdateItem/.test(request)).length, 1);
  assert.equal(requests.filter((request) => /SendItem/.test(request)).length, 1);
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
  assert.equal(result.lookback_years, 1);
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
  assert.deepEqual(_test.decodeCursor(result.next_cursor, "launch plan\u00001"), {
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

test("searchMessages supports a three-calendar-year lookback", async () => {
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
  const emptyFindItem = () => response(soap(`<m:FindItemResponse><m:ResponseMessages><m:FindItemResponseMessage ResponseClass="Success"><m:ResponseCode>NoError</m:ResponseCode><m:RootFolder IncludesLastItemInRange="true" TotalItemsInView="0"><t:Items/></m:RootFolder></m:FindItemResponseMessage></m:ResponseMessages></m:FindItemResponse>`));
  const responses = [
    response(soap(`<m:FindFolderResponse><m:ResponseMessages><m:FindFolderResponseMessage ResponseClass="Success"><m:ResponseCode>NoError</m:ResponseCode><m:RootFolder IncludesLastItemInRange="true" TotalItemsInView="1"><t:Folders><t:Folder><t:FolderId Id="archive-folder-id"/><t:DisplayName>Archive</t:DisplayName></t:Folder></t:Folders></m:RootFolder></m:FindFolderResponseMessage></m:ResponseMessages></m:FindFolderResponse>`)),
    emptyFindItem(),
    emptyFindItem(),
    emptyFindItem(),
  ];

  const result = await client.searchMessages({ query: "older project", lookbackYears: 3 });
  assert.equal(result.lookback_years, 3);
  assert.equal(result.searched_since, "2023-07-22T12:00:00.000Z");
  assert.equal(result.results.length, 0);
  assert.match(requests[1], /received:07\/22\/2023\.\.07\/22\/2026/);
  assert.match(requests[2], /sent:07\/22\/2023\.\.07\/22\/2026/);
});

test("searchMessages rejects lookbacks over three years", async () => {
  const client = clientWith([]);
  await assert.rejects(() => client.searchMessages({ query: "project", lookbackYears: 4 }), /1, 2, or 3 years/);
});

test("listAttachments returns safe attachment metadata", async () => {
  const requests = [];
  const client = clientWith([response(soap(`<m:GetItemResponse><m:ResponseMessages><m:GetItemResponseMessage ResponseClass="Success"><m:ResponseCode>NoError</m:ResponseCode><m:Items><t:Message><t:ItemId Id="message-id"/><t:Subject>Price history</t:Subject><t:Attachments>
    <t:FileAttachment><t:AttachmentId Id="file-id"/><t:Name>pricing.xlsx</t:Name><t:ContentType>application/vnd.openxmlformats-officedocument.spreadsheetml.sheet</t:ContentType><t:Size>2048</t:Size><t:IsInline>false</t:IsInline></t:FileAttachment>
    <t:ItemAttachment><t:AttachmentId Id="item-id"/><t:Name>Forwarded message</t:Name><t:Size>1024</t:Size></t:ItemAttachment>
  </t:Attachments></t:Message></m:Items></m:GetItemResponseMessage></m:ResponseMessages></m:GetItemResponse>`))], requests);
  const result = await client.listAttachments("message-id");
  assert.equal(result.attachment_count, 2);
  assert.deepEqual(result.attachments[0], {
    attachment_id: "file-id",
    name: "pricing.xlsx",
    content_type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    size: 2048,
    is_inline: false,
    kind: "file",
    reading_capability: "document_text",
    within_reading_limit: true,
  });
  assert.equal(result.attachments[1].reading_capability, "attached_email");
  assert.match(requests[0], /item:Attachments/);
  assert.equal(/CreateItem|UpdateItem|DeleteItem|MoveItem/.test(requests[0]), false);
});

test("readAttachment retrieves and extracts a text attachment without writes", async () => {
  const requests = [];
  const encoded = Buffer.from("Ball park price: $7.65\nMOQ: 15,000").toString("base64");
  const client = clientWith([
    response(soap(`<m:GetItemResponse><m:ResponseMessages><m:GetItemResponseMessage ResponseClass="Success"><m:ResponseCode>NoError</m:ResponseCode><m:Items><t:Message><t:Subject>Pricing</t:Subject><t:Attachments><t:FileAttachment><t:AttachmentId Id="attachment-id"/><t:Name>pricing.txt</t:Name><t:ContentType>text/plain</t:ContentType><t:Size>35</t:Size></t:FileAttachment></t:Attachments></t:Message></m:Items></m:GetItemResponseMessage></m:ResponseMessages></m:GetItemResponse>`)),
    response(soap(`<m:GetAttachmentResponse><m:ResponseMessages><m:GetAttachmentResponseMessage ResponseClass="Success"><m:ResponseCode>NoError</m:ResponseCode><m:Attachments><t:FileAttachment><t:AttachmentId Id="attachment-id"/><t:Name>pricing.txt</t:Name><t:ContentType>text/plain</t:ContentType><t:Content>${encoded}</t:Content></t:FileAttachment></m:Attachments></m:GetAttachmentResponseMessage></m:ResponseMessages></m:GetAttachmentResponse>`)),
  ], requests);
  const result = await client.readAttachment({ messageId: "message-id", attachmentId: "attachment-id" });
  assert.equal(result.capability, "text");
  assert.match(result.text, /\$7\.65/);
  assert.equal(result.truncated, false);
  assert.match(requests[1], /GetAttachment/);
  assert.equal(requests.some((request) => /CreateItem|UpdateItem|DeleteItem|MoveItem/.test(request)), false);
});

test("readAttachment expands an attached email", async () => {
  const client = clientWith([
    response(soap(`<m:GetItemResponse><m:ResponseMessages><m:GetItemResponseMessage ResponseClass="Success"><m:ResponseCode>NoError</m:ResponseCode><m:Items><t:Message><t:Subject>Fwd</t:Subject><t:Attachments><t:ItemAttachment><t:AttachmentId Id="item-id"/><t:Name>Original message</t:Name><t:Size>900</t:Size></t:ItemAttachment></t:Attachments></t:Message></m:Items></m:GetItemResponseMessage></m:ResponseMessages></m:GetItemResponse>`)),
    response(soap(`<m:GetAttachmentResponse><m:ResponseMessages><m:GetAttachmentResponseMessage ResponseClass="Success"><m:ResponseCode>NoError</m:ResponseCode><m:Attachments><t:ItemAttachment><t:AttachmentId Id="item-id"/><t:Name>Original message</t:Name><t:Item><t:Subject>Factory quote</t:Subject><t:DateTimeReceived>2026-07-01T10:00:00Z</t:DateTimeReceived><t:From><t:Mailbox><t:Name>Shirley</t:Name><t:EmailAddress>shirley@example.com</t:EmailAddress></t:Mailbox></t:From><t:ToRecipients><t:Mailbox><t:EmailAddress>etucker@metooshoes.com</t:EmailAddress></t:Mailbox></t:ToRecipients><t:Body BodyType="Text">Ball park pricing attached.</t:Body></t:Item></t:ItemAttachment></m:Attachments></m:GetAttachmentResponseMessage></m:ResponseMessages></m:GetAttachmentResponse>`)),
  ]);
  const result = await client.readAttachment({ messageId: "message-id", attachmentId: "item-id" });
  assert.equal(result.capability, "attached_email");
  assert.equal(result.attached_message.subject, "Factory quote");
  assert.equal(result.attached_message.sender.email, "shirley@example.com");
  assert.match(result.attached_message.body, /Ball park pricing/);
});

test("rejects non-Intermedia EWS URLs", () => {
  assert.throws(() => new EwsClient({
    url: "https://evil.example/EWS/Exchange.asmx",
    username: "u",
    password: "p",
    mailbox: "m@example.com",
  }), /serverdata\.net/);
});
