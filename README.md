# Me Too Shoes Intermedia Email Connectors

This repository contains two private, draft-only MCP connectors:

- `customer-service/` for `customerservice@metooshoes.com`
- `information/` for `information@metooshoes.com`

The root `render.yaml` deploys both services. Passwords are never stored in this repository; add each mailbox credential as a secret in Render.

Neither connector exposes a send-email tool.
