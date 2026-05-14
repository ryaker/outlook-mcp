# outlook-mcp

MCP server for Microsoft 365 (Outlook, OneDrive, Power Automate).

**Status: rebuild in progress (v3).** The v2 implementation has been wiped and is being rebuilt from scratch on TypeScript + the official MCP SDK, with a Streamable-HTTP transport so it can run as a Cowork-scheduled remote MCP. Earlier history is in `git log`.

## Stack

- Node 20+, TypeScript (strict), ESM
- `@modelcontextprotocol/sdk` — stdio + Streamable HTTP transports
- `zod` — tool input schemas + runtime validation
- `vitest` — tests
- Docker — for remote deployment

## Quick start

```bash
npm install
cp .env.example .env   # then fill in Azure credentials
npm run build
npm run start:stdio    # local dev
```

## Status of the rebuild

This scaffold contains only a `ping` smoke-test tool. Real M365 capabilities (auth, email, calendar, folders, rules, OneDrive, Power Automate, multi-account) will be added one module at a time via the design → implement → review pipeline.
