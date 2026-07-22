import test from "node:test";
import assert from "node:assert/strict";
import { attachmentCapability, extractAttachmentContent, MAX_ATTACHMENT_BYTES } from "../src/attachment-parser.js";

test("attachmentCapability recognizes common formats", () => {
  assert.equal(attachmentCapability({ name: "quote.pdf" }), "document_text");
  assert.equal(attachmentCapability({ name: "prices.xlsx" }), "document_text");
  assert.equal(attachmentCapability({ name: "notes.txt" }), "text");
  assert.equal(attachmentCapability({ name: "photo.jpg", contentType: "image/jpeg" }), "image");
  assert.equal(attachmentCapability({ name: "message", kind: "item" }), "attached_email");
  assert.equal(attachmentCapability({ name: "archive.zip" }), "unsupported");
});

test("extractAttachmentContent returns text and image data", async () => {
  const textResult = await extractAttachmentContent({
    name: "quote.txt",
    contentType: "text/plain",
    contentBase64: Buffer.from("Factory quote $7.65").toString("base64"),
  });
  assert.equal(textResult.text, "Factory quote $7.65");
  const imageBase64 = Buffer.from("fake-image").toString("base64");
  const imageResult = await extractAttachmentContent({ name: "photo.png", contentType: "image/png", contentBase64: imageBase64 });
  assert.equal(imageResult.image_base64, imageBase64);
});

test("extractAttachmentContent converts supported office documents", async () => {
  const parseCalls = [];
  const result = await extractAttachmentContent({
    name: "pricing.xlsx",
    contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    contentBase64: Buffer.from("fake-xlsx").toString("base64"),
  }, async (buffer, options) => {
    parseCalls.push({ buffer, options });
    return {
      warnings: [],
      to: async (format) => ({ value: "Style\tPrice\nGLEAM 02\t$7.65", format }),
    };
  });
  assert.match(result.text, /GLEAM 02/);
  assert.equal(parseCalls[0].options.fileType, "xlsx");
  assert.ok(parseCalls[0].options.abortSignal);
});

test("extractAttachmentContent rejects oversized attachments", async () => {
  const tooLarge = Buffer.alloc(MAX_ATTACHMENT_BYTES + 1).toString("base64");
  await assert.rejects(() => extractAttachmentContent({ name: "large.txt", contentType: "text/plain", contentBase64: tooLarge }), /10 MB/);
});
