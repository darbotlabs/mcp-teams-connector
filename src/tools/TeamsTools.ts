import { Client } from '@microsoft/microsoft-graph-client';

export class TeamsTools {
  constructor(private graphClient: Client) {}

  async sendTeamsMessage(params: {
    recipient: string;
    message: string;
    messageType?: 'chat' | 'channel';
  }): Promise<any> {
    try {
      console.log(`Sending ${params.messageType || 'chat'} message to: ${params.recipient}`);
      
      if (params.messageType === 'channel' || params.recipient.includes('/channels/')) {
        // Send to channel
        return await this.sendChannelMessage(params.recipient, params.message);
      } else {
        // Send to chat
        return await this.sendChatMessage(params.recipient, params.message);
      }
    } catch (error: any) {
      console.error('Failed to send Teams message:', error);
      return {
        content: [{
          type: 'text',
          text: `Failed to send message: ${error.message || 'Unknown error'}`
        }],
        isError: true
      };
    }
  }

  private async sendChatMessage(recipient: string, message: string): Promise<any> {
    try {
      // First, find or create a chat with the recipient
      const chat = await this.findOrCreateChat(recipient);
      
      // Send message to the chat
      const chatMessage = {
        body: {
          content: message,
          contentType: 'text'
        }
      };

      const sentMessage = await this.graphClient
        .api(`/chats/${chat.id}/messages`)
        .post(chatMessage);

      return {
        content: [{
          type: 'text',
          text: `Message sent successfully to ${recipient}\nMessage ID: ${sentMessage.id}`
        }]
      };
    } catch (error: any) {
      console.error('Failed to send chat message:', error);
      throw new Error(`Failed to send chat message: ${error.message}`);
    }
  }

  private async sendChannelMessage(channelPath: string, message: string): Promise<any> {
    try {
      const chatMessage = {
        body: {
          content: message,
          contentType: 'text'
        }
      };

      const sentMessage = await this.graphClient
        .api(`${channelPath}/messages`)
        .post(chatMessage);

      return {
        content: [{
          type: 'text',
          text: `Message sent to channel successfully\nMessage ID: ${sentMessage.id}`
        }]
      };
    } catch (error: any) {
      console.error('Failed to send channel message:', error);
      throw new Error(`Failed to send channel message: ${error.message}`);
    }
  }

  private async findOrCreateChat(userEmail: string): Promise<any> {
    try {
      // Get user ID from email
      const user = await this.graphClient
        .api(`/users/${userEmail}`)
        .select('id,displayName')
        .get();

      // Try to find existing chat
      const chats = await this.graphClient
        .api('/me/chats')
        .filter(`chatType eq 'oneOnOne'`)
        .expand('members')
        .get();

      // Look for existing chat with this user
      for (const chat of chats.value) {
        const memberIds = chat.members?.map((m: any) => m.userId) || [];
        if (memberIds.includes(user.id)) {
          console.log(`Found existing chat with ${userEmail}`);
          return chat;
        }
      }

      // Create new chat if not found
      console.log(`Creating new chat with ${userEmail}`);
      const newChat = {
        chatType: 'oneOnOne',
        members: [
          {
            '@odata.type': '#microsoft.graph.aadUserConversationMember',
            roles: ['owner'],
            'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${user.id}')`
          }
        ]
      };

      return await this.graphClient
        .api('/chats')
        .post(newChat);
    } catch (error: any) {
      console.error('Failed to find or create chat:', error);
      throw new Error(`Failed to find or create chat with ${userEmail}: ${error.message}`);
    }
  }

  async listTeams(): Promise<any> {
    try {
      console.log('Retrieving user\'s Teams');
      
      const teams = await this.graphClient
        .api('/me/joinedTeams')
        .select('id,displayName,description,webUrl')
        .get();

      if (!teams.value || teams.value.length === 0) {
        return {
          content: [{
            type: 'text',
            text: 'You are not a member of any Teams.'
          }]
        };
      }

      const teamList = teams.value.map((team: any, index: number) => {
        let text = `${index + 1}. **${team.displayName}**`;
        text += `\n   ID: ${team.id}`;
        
        if (team.description) {
          text += `\n   Description: ${team.description}`;
        }
        
        if (team.webUrl) {
          text += `\n   Web URL: ${team.webUrl}`;
        }
        
        return text;
      }).join('\n\n');

      return {
        content: [{
          type: 'text',
          text: `Your Teams (${teams.value.length} total):\n\n${teamList}`
        }]
      };
    } catch (error: any) {
      console.error('Failed to list teams:', error);
      return {
        content: [{
          type: 'text',
          text: `Failed to list teams: ${error.message || 'Unknown error'}`
        }],
        isError: true
      };
    }
  }

  async listChannels(teamId: string): Promise<any> {
    try {
      console.log(`Retrieving channels for team: ${teamId}`);
      
      const channels = await this.graphClient
        .api(`/teams/${teamId}/channels`)
        .select('id,displayName,description,webUrl,membershipType')
        .get();

      if (!channels.value || channels.value.length === 0) {
        return {
          content: [{
            type: 'text',
            text: `No channels found for this team.`
          }]
        };
      }

      const channelList = channels.value.map((channel: any, index: number) => {
        let text = `${index + 1}. **${channel.displayName}**`;
        text += `\n   ID: ${channel.id}`;
        text += `\n   Type: ${channel.membershipType || 'standard'}`;
        
        if (channel.description) {
          text += `\n   Description: ${channel.description}`;
        }
        
        if (channel.webUrl) {
          text += `\n   Web URL: ${channel.webUrl}`;
        }
        
        // Add channel path for sending messages
        text += `\n   Channel Path: /teams/${teamId}/channels/${channel.id}`;
        
        return text;
      }).join('\n\n');

      return {
        content: [{
          type: 'text',
          text: `Channels in team (${channels.value.length} total):\n\n${channelList}`
        }]
      };
    } catch (error: any) {
      console.error('Failed to list channels:', error);
      return {
        content: [{
          type: 'text',
          text: `Failed to list channels: ${error.message || 'Unknown error'}`
        }],
        isError: true
      };
    }
  }

  async getRecentChats(): Promise<any> {
    try {
      console.log('Retrieving recent chats');
      
      const chats = await this.graphClient
        .api('/me/chats')
        .orderby('lastUpdatedDateTime desc')
        .top(20)
        .expand('members')
        .get();

      if (!chats.value || chats.value.length === 0) {
        return {
          content: [{
            type: 'text',
            text: 'No recent chats found.'
          }]
        };
      }

      const chatList = chats.value.map((chat: any, index: number) => {
        let text = `${index + 1}. **${chat.chatType}** chat`;
        text += `\n   ID: ${chat.id}`;
        
        if (chat.topic) {
          text += `\n   Topic: ${chat.topic}`;
        }
        
        if (chat.members && chat.members.length > 0) {
          const memberNames = chat.members
            .filter((m: any) => m.displayName)
            .map((m: any) => m.displayName)
            .join(', ');
          
          if (memberNames) {
            text += `\n   Members: ${memberNames}`;
          }
        }
        
        if (chat.lastUpdatedDateTime) {
          const lastUpdated = new Date(chat.lastUpdatedDateTime);
          text += `\n   Last Updated: ${lastUpdated.toLocaleString()}`;
        }
        
        return text;
      }).join('\n\n');

      return {
        content: [{
          type: 'text',
          text: `Recent chats (${chats.value.length} total):\n\n${chatList}`
        }]
      };
    } catch (error: any) {
      console.error('Failed to get recent chats:', error);
      return {
        content: [{
          type: 'text',
          text: `Failed to get recent chats: ${error.message || 'Unknown error'}`
        }],
        isError: true
      };
    }
  }
}