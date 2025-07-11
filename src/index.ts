import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
  Tool,
  CallToolResult,
  ListToolsResult,
  TextContent
} from '@modelcontextprotocol/sdk/types.js';
import { GraphAuthProvider } from './auth/GraphAuthProvider.js';
import { CalendarTools } from './tools/CalendarTools.js';
import { TeamsTools } from './tools/TeamsTools.js';

class MCPTeamsServer {
  private server: Server;
  private authProvider: GraphAuthProvider;
  private calendarTools: CalendarTools | null = null;
  private teamsTools: TeamsTools | null = null;

  constructor() {
    this.authProvider = new GraphAuthProvider();
    this.server = new Server({
      name: 'mcp-teams-connector',
      version: '1.0.0',
      description: 'MCP connector for Microsoft Teams and Outlook integration'
    }, {
      capabilities: {
        tools: {}
      }
    });

    this.setupHandlers();
  }

  private setupHandlers() {
    // Tools call handler
    this.server.setRequestHandler(CallToolRequestSchema, async (request) => {
      const { name, arguments: args } = request.params;

      if (!this.calendarTools || !this.teamsTools) {
        return {
          content: [{
            type: 'text',
            text: 'Server not properly initialized. Please restart the MCP server.'
          }],
          isError: true
        };
      }

      try {
        switch (name) {
          case 'createMeeting':
            return await this.calendarTools.createMeeting(args as any);
          
          case 'findAvailability':
            return await this.calendarTools.findAvailability(args as any);
          
          case 'listUpcomingMeetings':
            return await this.calendarTools.listUpcomingMeetings(args as any);
          
          case 'sendTeamsMessage':
            return await this.teamsTools.sendTeamsMessage(args as any);
          
          case 'listTeams':
            return await this.teamsTools.listTeams();
          
          default:
            return {
              content: [{
                type: 'text',
                text: `Unknown tool: ${name}`
              }],
              isError: true
            };
        }
      } catch (error) {
        console.error(`Error executing tool ${name}:`, error);
        return {
          content: [{
            type: 'text',
            text: `Error executing ${name}: ${error instanceof Error ? error.message : 'Unknown error'}`
          }],
          isError: true
        };
      }
    });

    // List available tools
    this.server.setRequestHandler(ListToolsRequestSchema, async () => {
      return {
        tools: [
          {
            name: 'createMeeting',
            description: 'Creates a new meeting in Microsoft Teams or Outlook calendar',
            inputSchema: {
              type: 'object',
              properties: {
                subject: { 
                  type: 'string', 
                  description: 'Meeting title' 
                },
                startTime: { 
                  type: 'string', 
                  description: 'ISO format start time (e.g., 2025-07-11T14:00:00Z)' 
                },
                endTime: { 
                  type: 'string', 
                  description: 'ISO format end time (e.g., 2025-07-11T15:00:00Z)' 
                },
                attendees: { 
                  type: 'array', 
                  items: { type: 'string' },
                  description: 'Email addresses of attendees'
                },
                body: { 
                  type: 'string', 
                  description: 'Meeting description or agenda' 
                },
                location: { 
                  type: 'string', 
                  description: 'Meeting location (room, address, or virtual)' 
                },
                isOnline: { 
                  type: 'boolean', 
                  description: 'Create as Teams online meeting',
                  default: true
                }
              },
              required: ['subject', 'startTime', 'endTime']
            }
          },
          {
            name: 'findAvailability',
            description: 'Check availability for multiple attendees and suggest meeting times',
            inputSchema: {
              type: 'object',
              properties: {
                attendees: { 
                  type: 'array', 
                  items: { type: 'string' },
                  description: 'Email addresses to check availability for'
                },
                startDate: { 
                  type: 'string', 
                  description: 'Start date for availability search (ISO format)'
                },
                endDate: { 
                  type: 'string', 
                  description: 'End date for availability search (ISO format)'
                },
                duration: { 
                  type: 'number', 
                  description: 'Meeting duration in minutes'
                }
              },
              required: ['attendees', 'startDate', 'endDate', 'duration']
            }
          },
          {
            name: 'listUpcomingMeetings',
            description: 'Get upcoming meetings from calendar',
            inputSchema: {
              type: 'object',
              properties: {
                days: { 
                  type: 'number', 
                  default: 7,
                  description: 'Number of days to look ahead'
                },
                includeDetails: { 
                  type: 'boolean', 
                  default: false,
                  description: 'Include detailed meeting information'
                }
              }
            }
          },
          {
            name: 'sendTeamsMessage',
            description: 'Send a message to a Teams channel or chat',
            inputSchema: {
              type: 'object',
              properties: {
                recipient: { 
                  type: 'string', 
                  description: 'Email address for chat or channel path for channel messages'
                },
                message: { 
                  type: 'string', 
                  description: 'Message content to send'
                },
                messageType: { 
                  type: 'string', 
                  enum: ['chat', 'channel'],
                  description: 'Type of message (chat for direct message, channel for channel message)',
                  default: 'chat'
                }
              },
              required: ['recipient', 'message']
            }
          },
          {
            name: 'listTeams',
            description: 'List all Teams that the user is a member of',
            inputSchema: {
              type: 'object',
              properties: {}
            }
          }
        ]
      };
    });

    // Server initialization is handled automatically by the MCP SDK
  }

  async start() {
    try {
      // Authenticate before starting server
      console.log('MCP Teams Connector starting...');
      console.log('Authenticating with Microsoft Graph...');
      
      await this.authProvider.authenticate();
      
      // Validate tenant and user
      const isValidTenant = await this.authProvider.validateTenant();
      if (!isValidTenant) {
        throw new Error('Invalid tenant. This connector is configured for Cypherdyne tenant only.');
      }

      const isValidUser = await this.authProvider.validateUser();
      if (!isValidUser) {
        throw new Error('Invalid user. This connector is configured for @timelarp.com users only.');
      }

      // Initialize tools with authenticated Graph client
      const graphClient = await this.authProvider.getGraphClient();
      this.calendarTools = new CalendarTools(graphClient);
      this.teamsTools = new TeamsTools(graphClient);

      console.log('Authentication successful!');
      console.log(`Authenticated as: ${this.authProvider.getCurrentUser()?.username}`);

      // Start MCP server
      const transport = new StdioServerTransport();
      await this.server.connect(transport);
      
      console.log('MCP Teams Connector is running and ready to accept requests...');
      
    } catch (error) {
      console.error('Failed to start MCP Teams Connector:', error);
      process.exit(1);
    }
  }

  async stop() {
    try {
      console.log('Shutting down MCP Teams Connector...');
      await this.authProvider.signOut();
      console.log('MCP Teams Connector stopped.');
    } catch (error) {
      console.error('Error during shutdown:', error);
    }
  }
}

// Handle graceful shutdown
process.on('SIGINT', async () => {
  console.log('\nReceived SIGINT, shutting down gracefully...');
  if (server) {
    await server.stop();
  }
  process.exit(0);
});

process.on('SIGTERM', async () => {
  console.log('\nReceived SIGTERM, shutting down gracefully...');
  if (server) {
    await server.stop();
  }
  process.exit(0);
});

// Start the server
const server = new MCPTeamsServer();
server.start().catch((error) => {
  console.error('Failed to start server:', error);
  process.exit(1);
});