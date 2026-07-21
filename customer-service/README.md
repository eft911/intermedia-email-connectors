# Intermedia Exchange Connector

Private MCP connector for an Intermedia Hosted Exchange customer-service mailbox. It can read recent mail and save plain-text reply drafts. It has no send-email capability.

## Tools

- `intermedia_health_check` — verifies mailbox access and returns inbox counts.
- `list_customer_messages` — lists recent/unread messages, excluding previously drafted messages by default.
- `get_customer_message` — retrieves one message as plain text.
- `create_customer_reply_draft` — saves a reply draft and adds the `GPT Drafted` category to the source message.

## Security design

- EWS credentials are read only from deployment environment variables.
- The EWS endpoint must be HTTPS, end in `/EWS/Exchange.asmx`, and use a `serverdata.net` hostname.
- The public `/mcp` route returns 404. MCP is exposed only at `/mcp/<MCP_PATH_SECRET>`.
- There is deliberately no send tool.
- Use a dedicated, least-privileged mailbox credential where possible. Do not commit `.env` or credentials.

## Required Intermedia details

In HostPilot, open **Home → Exchange servers and settings → Exchange Proxy Setting**. Build the EWS URL exactly as Intermedia specifies:

`https://east.exch028.serverdata.net/EWS/Exchange.asmx`

The account in `EWS_USERNAME` must either own `SHARED_MAILBOX` or have Full Access to it.

## Local setup

```bash
cp .env.example .env
npm install
```

Load the `.env` values in your shell or deployment platform, then run:

```bash
npm start
```

Health endpoint: `http://localhost:3000/health`

MCP endpoint: `http://localhost:3000/mcp/<MCP_PATH_SECRET>`

## Render deployment

1. Push this folder to a private GitHub repository.
2. In Render, create a Blueprint from `render.yaml`.
3. Add `EWS_URL`, `EWS_USERNAME`, and `EWS_PASSWORD` as secret environment values.
4. Confirm `SHARED_MAILBOX=customerservice@metooshoes.com`.
5. Copy the generated `MCP_PATH_SECRET` without exposing it publicly.
6. Connect ChatGPT to `https://<render-host>/mcp/<MCP_PATH_SECRET>`.

Intermedia Exchange 2FA may require first-device approval. If EWS returns 401, verify the credentials and Exchange 2FA/device authorization in HostPilot.

## Hourly automation prompt

After the connector is deployed and tested, use this prompt for an hourly automation:

> Check the configured customer-service inbox for recent customer inquiries that do not already have a GPT Drafted category. Include read and unread messages so a staff preview does not suppress drafting. Skip messages that already received a human reply, plus spam, newsletters, receipts, and automated notices. Read each qualifying message and save a concise, warm customer-service reply draft. Never send email. Do not invent order status, refunds, delivery dates, inventory, policies, or commitments; when facts are missing, ask for the order number or state what must be verified. If no messages qualify, take no action.
