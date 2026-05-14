import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";
import { z } from "zod";

const SERVER_NAME = "outlook-mcp";
const SERVER_VERSION = "3.0.0";

const PingInput = z.object({
  message: z.string().optional(),
});

export function createServer(): Server {
  const server = new Server(
    { name: SERVER_NAME, version: SERVER_VERSION },
    { capabilities: { tools: {} } },
  );

  server.setRequestHandler(ListToolsRequestSchema, async () => ({
    tools: [
      {
        name: "ping",
        description:
          "Smoke-test tool. Echoes back the optional message so we can verify the transport is alive.",
        inputSchema: {
          type: "object",
          properties: {
            message: { type: "string" },
          },
        },
      },
    ],
  }));

  server.setRequestHandler(CallToolRequestSchema, async (request) => {
    if (request.params.name === "ping") {
      const { message } = PingInput.parse(request.params.arguments ?? {});
      return {
        content: [
          {
            type: "text",
            text: `pong${message ? `: ${message}` : ""}`,
          },
        ],
      };
    }
    throw new Error(`Unknown tool: ${request.params.name}`);
  });

  return server;
}
