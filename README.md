# MCP Teams Connector

A standalone Model Context Protocol (MCP) server that bridges Microsoft Teams and Outlook with AI tools like Claude Desktop and VS Code. This connector enables AI assistants to manage calendar events, send Teams messages, and handle meetings through Microsoft Graph API.

## Features

- 🗓️ **Calendar Management**: Create, update, and list meetings
- 🔍 **Availability Checking**: Find optimal meeting times for multiple attendees
- 💬 **Teams Messaging**: Send messages to Teams channels and direct chats
- 🔐 **Secure Authentication**: OAuth2 with credential caching and keychain storage
- 🏢 **Tenant Validation**: Configured for Cypherdyne tenant security

## Available Tools

### Calendar Tools
- `createMeeting` - Create new meetings with Teams integration
- `findAvailability` - Check attendee availability and suggest optimal times
- `listUpcomingMeetings` - Retrieve upcoming calendar events

### Teams Tools
- `sendTeamsMessage` - Send messages to Teams chats or channels
- `listTeams` - List all Teams the user is a member of

## Prerequisites

- Node.js 18.0.0 or higher
- Microsoft 365 account with Teams and Outlook access
- Access to Cypherdyne tenant (`@timelarp.com` domain)

## Installation

### 1. Clone and Install Dependencies

```bash
git clone https://github.com/darbotron/mcp-teams-connector.git
cd mcp-teams-connector
npm install
```

### 2. Build the Project

```bash
npm run build
```

### 3. Configure MCP Client

#### For Claude Desktop

1. Locate your Claude Desktop configuration file:
   - Windows: `%APPDATA%\Claude\claude_desktop_config.json`
   - macOS: `~/Library/Application Support/Claude/claude_desktop_config.json`

2. Add the MCP server configuration:

```json
{
  "mcpServers": {
    "mcp-teams-connector": {
      "command": "node",
      "args": ["C:/path/to/mcp-teams-connector/dist/index.js"],
      "env": {
        "NODE_ENV": "production"
      }
    }
  }
}
```

3. Replace `C:/path/to/mcp-teams-connector` with the actual path to your installation

4. Restart Claude Desktop

#### For VS Code

1. Install the MCP extension for VS Code

2. Add configuration to your workspace settings (`.vscode/settings.json`):

```json
{
  "mcp.servers": {
    "teams-connector": {
      "command": "node",
      "args": ["${workspaceFolder}/mcp-teams-connector/dist/index.js"],
      "environment": {
        "NODE_ENV": "production"
      }
    }
  }
}
```

3. Reload VS Code window

## First Run & Authentication

When you first use the MCP server, it will:

1. Open your default browser for Microsoft authentication
2. Prompt you to sign in with your `@timelarp.com` account
3. Request necessary permissions for Calendar and Teams access
4. Cache authentication tokens securely

The authentication process only needs to be completed once. Tokens are automatically refreshed as needed.

## Usage Examples

Once configured, you can use natural language commands with your AI assistant:

### Calendar Management

```
"Create a Teams meeting tomorrow at 2 PM with john@company.com about Project Review"

"Schedule a 1-hour meeting next Friday at 10 AM with the engineering team"

"Check my availability next week for a 30-minute meeting"

"List my meetings for the next 3 days with full details"
```

### Teams Messaging

```
"Send a Teams message to sarah@company.com saying 'The report is ready for review'"

"Message the development team channel about the deployment status"

"List all the Teams I'm a member of"
```

### Availability Checking

```
"Find available time slots for a 60-minute meeting with john@company.com and jane@company.com next week"

"When are all attendees free for a 2-hour workshop between Monday and Wednesday?"
```

## Development

### Development Mode

```bash
npm run dev
```

This starts the server with hot reloading using `tsx watch`.

### Project Structure

```
mcp-teams-connector/
├── src/
│   ├── auth/
│   │   └── GraphAuthProvider.ts    # Microsoft Graph authentication
│   ├── tools/
│   │   ├── CalendarTools.ts        # Calendar management
│   │   └── TeamsTools.ts           # Teams messaging
│   └── index.ts                    # Main MCP server
├── config/
│   ├── mcp-settings.json           # Claude Desktop config template
│   └── vscode-settings.json        # VS Code config template
├── dist/                           # Compiled JavaScript
└── .cache/                         # Authentication token cache
```

### Available Scripts

- `npm start` - Run the compiled server
- `npm run dev` - Development mode with hot reloading
- `npm run build` - Build the TypeScript project
- `npm run clean` - Clean the dist directory

## Security & Privacy

- **Tenant Validation**: Only allows users from the Cypherdyne tenant
- **Secure Token Storage**: Uses OS keychain for refresh token storage
- **Local Operation**: Runs entirely on your machine, no data sent to third parties
- **Minimal Permissions**: Requests only necessary Microsoft Graph API scopes

## Troubleshooting

### Authentication Issues

1. **Browser doesn't open**: Manual authentication URL will be displayed in the console
2. **Permission denied**: Ensure your account has necessary Teams and Calendar permissions
3. **Token expired**: Delete `.cache/tokens.json` and restart to re-authenticate

### Connection Issues

1. **Port 3000 in use**: The authentication process uses port 3000 temporarily
2. **Firewall blocking**: Ensure localhost:3000 is accessible during authentication

### MCP Client Issues

1. **Claude Desktop**: Check the configuration file path and syntax
2. **VS Code**: Ensure the MCP extension is properly installed
3. **Path issues**: Verify the absolute path to `dist/index.js` is correct

## Configuration

### Environment Variables

The server uses the following configuration from `darbot_config.md`:

- **Tenant ID**: `6b104499-c49f-45dc-b3a2-df95efd6eeb4` (Cypherdyne)
- **Client ID**: `bedaebf0-4f7a-4c5b-8861-e082001a8193` (MeetingAssist app)
- **Authority**: `https://login.microsoftonline.com/6b104499-c49f-45dc-b3a2-df95efd6eeb4`

### Supported Scopes

- `User.Read` - Read user profile
- `Calendars.ReadWrite` - Manage calendar events
- `Mail.Send` - Send emails (for meeting invitations)
- `Chat.ReadWrite` - Access Teams chats
- `ChannelMessage.Send` - Send channel messages
- `Team.ReadBasic.All` - List user's teams
- `OnlineMeetings.ReadWrite` - Create Teams meetings

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests if applicable
5. Submit a pull request

## License

MIT License - see LICENSE file for details

## Support

For issues and questions:
1. Check the troubleshooting section above
2. Review the authentication logs in the console
3. Verify your Microsoft 365 permissions
4. Ensure you're using a `@timelarp.com` account

## Version History

- **1.0.0** - Initial release with calendar and Teams integration
  - Microsoft Graph authentication with caching
  - Calendar event management
  - Teams messaging capabilities
  - MCP protocol implementation