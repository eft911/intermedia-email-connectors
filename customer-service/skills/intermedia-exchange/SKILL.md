---
name: intermedia-exchange
description: Use when the user wants to inspect the configured Intermedia customer-service inbox or create safe reply drafts without sending email.
---

# Intermedia Exchange Customer Service

Use the bundled MCP tools for the configured Intermedia Hosted Exchange mailbox.

## Workflow

1. Call `intermedia_health_check` before the first mailbox operation in a thread.
2. Use `list_customer_messages` to identify unread, recent messages that do not already have a generated draft.
3. Use `get_customer_message` before drafting so the response is grounded in the full message.
4. Draft in a warm, concise customer-service tone. Do not invent order status, refunds, delivery dates, inventory, policies, or commitments.
5. Call `create_customer_reply_draft` only when the user asked for a draft or an approved automation is running.
6. Never claim a message was sent. This plugin intentionally has no send tool.

## Automated Drafting Rules

- Skip spam, newsletters, receipts, delivery-system notices, and automated messages.
- Draft only when a customer appears to need a response.
- If order facts are missing, ask for the order number or state what must be verified.
- Keep sensitive internal notes out of the draft body.
- The connector adds `GPT Drafted` to the source message after saving a draft to prevent duplicates.
