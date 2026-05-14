#!/usr/bin/env node
import "dotenv/config";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { createServer } from "./server.js";

type Transport = "stdio" | "http";

function resolveTransport(): Transport {
  const raw = process.env.MCP_TRANSPORT?.toLowerCase() ?? "stdio";
  if (raw === "stdio" || raw === "http") return raw;
  throw new Error(
    `Invalid MCP_TRANSPORT: "${raw}". Expected "stdio" or "http".`,
  );
}

async function main(): Promise<void> {
  const transport = resolveTransport();
  const server = createServer();

  if (transport === "stdio") {
    await server.connect(new StdioServerTransport());
    console.error("outlook-mcp connected on stdio");
    return;
  }

  // Streamable HTTP transport — to be wired in the design phase
  throw new Error(
    "HTTP transport not yet implemented. Set MCP_TRANSPORT=stdio for now.",
  );
}

main().catch((error) => {
  console.error("Fatal:", error);
  process.exit(1);
});
