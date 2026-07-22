import { OfficeParser } from "officeparser";

export const MAX_ATTACHMENT_BYTES = 10 * 1024 * 1024;
export const MAX_EXTRACTED_CHARS = 100_000;

const OFFICE_TYPES = new Set([
  "pdf", "docx", "pptx", "xlsx", "odt", "odp", "ods", "rtf", "csv", "md", "html", "epub",
]);
const TEXT_EXTENSIONS = new Set([
  "txt", "text", "log", "json", "xml", "yaml", "yml", "ini", "cfg", "sql", "js", "ts", "css", "ics", "vcf", "eml",
]);

function extension(name) {
  const match = String(name || "").toLocaleLowerCase().match(/\.([a-z0-9]+)$/);
  return match?.[1] || "";
}

function normalizeText(value) {
  return String(value || "").replaceAll("\u0000", "").replace(/\r\n/g, "\n").trim();
}

function truncateText(value) {
  const text = normalizeText(value);
  return {
    text: text.slice(0, MAX_EXTRACTED_CHARS),
    truncated: text.length > MAX_EXTRACTED_CHARS,
    extracted_characters: Math.min(text.length, MAX_EXTRACTED_CHARS),
  };
}

export function attachmentCapability({ name, contentType, kind = "file" }) {
  if (kind === "item") return "attached_email";
  const ext = extension(name);
  const mime = String(contentType || "").toLocaleLowerCase();
  if (mime.startsWith("image/") || ["png", "jpg", "jpeg", "gif", "webp", "bmp", "tif", "tiff"].includes(ext)) return "image";
  if (mime.startsWith("text/") || TEXT_EXTENSIONS.has(ext)) return "text";
  if (OFFICE_TYPES.has(ext)) return "document_text";
  return "unsupported";
}

export async function extractAttachmentContent({ name, contentType, contentBase64, kind = "file" }, parseOffice = OfficeParser.parseOffice) {
  const buffer = Buffer.from(contentBase64 || "", "base64");
  if (buffer.length > MAX_ATTACHMENT_BYTES) {
    throw new Error(`Attachment exceeds the ${MAX_ATTACHMENT_BYTES / 1024 / 1024} MB reading limit.`);
  }
  const capability = attachmentCapability({ name, contentType, kind });
  const base = {
    name,
    content_type: contentType || "application/octet-stream",
    size: buffer.length,
    capability,
  };
  if (capability === "image") {
    return { ...base, image_base64: contentBase64, text: null, truncated: false };
  }
  if (capability === "text") {
    return { ...base, ...truncateText(buffer.toString("utf8")) };
  }
  if (capability === "document_text") {
    const fileType = extension(name);
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), 25_000);
    try {
      const ast = await parseOffice(buffer, { fileType, abortSignal: controller.signal });
      const conversion = await ast.to("text", { includeImages: false });
      return { ...base, ...truncateText(conversion.value), warnings: (ast.warnings || []).map((warning) => warning.message || String(warning)) };
    } catch (error) {
      const message = error?.name === "AbortError"
        ? "Attachment text extraction timed out."
        : "Attachment text could not be extracted. It may be password-protected, damaged, or unsupported.";
      return { ...base, text: null, truncated: false, extraction_error: message };
    } finally {
      clearTimeout(timer);
    }
  }
  return {
    ...base,
    text: null,
    truncated: false,
    extraction_error: "This attachment format is not supported for safe text extraction.",
  };
}

export const _test = { extension, normalizeText, truncateText };
