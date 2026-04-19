# M365 Assistant for Claude Desktop on macOS

This guide is for Mac users who want Claude Desktop to use the `m365-assistant` MCP server for Outlook, calendar, OneDrive, and related Microsoft 365 tools.

## Before You Start

You need these things from Raf:

- The GitHub repository URL
- These Azure values:
  - `MS_CLIENT_ID`
  - `MS_CLIENT_SECRET`
  - `MS_TENANT_ID`
- Confirmation that you should use the shared Azure app registration

You do not need anyone else's token file. You will generate your own local token during setup.

## What You Need Installed

- Claude Desktop
- Git
- Node.js 18 or newer
- A Claude account

Local MCP servers in Claude Desktop work on free Claude accounts, so a paid Claude subscription is not required.

## Quick Setup

Open Terminal and run these commands:

```bash
node -v || brew install node
mkdir -p ~/code
cd ~/code
git clone https://github.com/ryaker/outlook-mcp.git
cd outlook-mcp
npm ci
cp .env.example .env
pwd
```

If Raf gives you a different repo URL, use that instead of `https://github.com/ryaker/outlook-mcp.git`.

The last command prints your full repo path. Keep that handy because you will paste it into the Claude Desktop config in the next step.

## Fill In `.env`

Open the project environment file:

```bash
open -e .env
```

Replace the placeholder values with the Azure values from Raf:

```env
MS_CLIENT_ID=your-client-id
MS_CLIENT_SECRET=your-client-secret-value
MS_TENANT_ID=your-tenant-id
USE_TEST_MODE=false
```

Important:

- Use the client secret value, not the secret ID
- Keep `.env` private
- Do not commit `.env` to git

## Add It To Claude Desktop

Claude Desktop on macOS uses this config file:

```text
~/Library/Application Support/Claude/claude_desktop_config.json
```

Open it:

```bash
open -e ~/Library/Application\ Support/Claude/claude_desktop_config.json
```

If the file does not exist yet, create it first:

```bash
mkdir -p ~/Library/Application\ Support/Claude
touch ~/Library/Application\ Support/Claude/claude_desktop_config.json
open -e ~/Library/Application\ Support/Claude/claude_desktop_config.json
```

Paste this JSON and replace the path with your actual repo path from `pwd`:

```json
{
  "mcpServers": {
    "m365-assistant": {
      "command": "node",
      "args": [
        "/Users/YOUR_USERNAME/code/outlook-mcp/index.js"
      ],
      "env": {
        "USE_TEST_MODE": "false"
      }
    }
  }
}
```

Example:

```json
{
  "mcpServers": {
    "m365-assistant": {
      "command": "node",
      "args": [
        "/Users/jane/code/outlook-mcp/index.js"
      ],
      "env": {
        "USE_TEST_MODE": "false"
      }
    }
  }
}
```

If you already have other MCP servers in that file, keep them and only add the `m365-assistant` block without breaking the JSON.

## Start The Auth Server

In Terminal, run:

```bash
cd ~/code/outlook-mcp
npm run auth-server
```

Leave that Terminal window open.

## Restart Claude Desktop

Fully quit Claude Desktop, then open it again.

## Authenticate

In Claude Desktop, ask Claude:

```text
Use m365-assistant to authenticate my Microsoft account.
```

Claude should return a Microsoft sign-in URL. Open it, sign in, and finish the auth flow.

Your local token file will be saved here:

```text
~/.outlook-mcp-tokens.json
```

That file belongs only to your machine and account.

## Test It

Try one of these prompts in Claude Desktop:

- `Use m365-assistant to list my next 5 calendar events.`
- `Use m365-assistant to show my last 10 inbox emails.`
- `Use m365-assistant to check my authentication status.`

If Claude calls the tool successfully, setup is complete.

## Troubleshooting

### Claude does not show the server

- Make sure `~/Library/Application Support/Claude/claude_desktop_config.json` contains valid JSON
- Make sure the `index.js` path is correct
- Fully quit and reopen Claude Desktop

### `node: command not found`

Run:

```bash
brew install node
```

### `Cannot find module`

From the repo folder, run:

```bash
npm ci
```

If that fails, try:

```bash
npm install
```

### Authentication fails

Check:

- `.env` contains the correct Azure values
- The redirect URI in Azure is `http://localhost:3333/auth/callback`
- `npm run auth-server` is still running
- You used the client secret value, not the secret ID

### Port 3333 is already in use

Run:

```bash
lsof -i :3333
```

Then stop the conflicting process and rerun:

```bash
npm run auth-server
```

## What Raf Needs To Do

Usually Raf only needs to do these things once:

- Keep the GitHub repo accessible
- Decide whether coworkers should use one shared Azure app registration or separate ones
- Share the approved Azure values securely
- Tell coworkers which repo URL to clone

If those are already in place, there is nothing else Raf needs to do per Mac user.
