import { PublicClientApplication, AccountInfo, AuthenticationResult } from '@azure/msal-node';
import { Client } from '@microsoft/microsoft-graph-client';
import * as keytar from 'keytar';
import * as fs from 'fs/promises';
import * as path from 'path';
import open from 'open';
import express from 'express';

const SERVICE_NAME = 'MCPTeamsConnector';
const CACHE_DIR = path.join(process.cwd(), '.cache');
const TOKEN_CACHE_FILE = path.join(CACHE_DIR, 'tokens.json');

export class GraphAuthProvider {
  private msalClient: PublicClientApplication;
  private graphClient: Client | null = null;
  private currentAccount: AccountInfo | null = null;
  
  constructor() {
    // Using the MeetingAssist app registration from darbot_config
    const msalConfig = {
      auth: {
        clientId: 'bedaebf0-4f7a-4c5b-8861-e082001a8193', // MeetingAssist app
        authority: 'https://login.microsoftonline.com/6b104499-c49f-45dc-b3a2-df95efd6eeb4',
        redirectUri: 'http://localhost:3000/redirect'
      },
      cache: {
        cachePlugin: {
          beforeCacheAccess: async (cacheContext: any) => {
            try {
              const data = await fs.readFile(TOKEN_CACHE_FILE, 'utf-8');
              cacheContext.tokenCache.deserialize(data);
            } catch (error) {
              // Cache doesn't exist yet, this is normal on first run
              console.log('No existing token cache found, will create new one');
            }
          },
          afterCacheAccess: async (cacheContext: any) => {
            if (cacheContext.hasChanged) {
              try {
                await fs.mkdir(CACHE_DIR, { recursive: true });
                await fs.writeFile(TOKEN_CACHE_FILE, cacheContext.tokenCache.serialize());
                console.log('Token cache updated');
              } catch (error) {
                console.error('Failed to write token cache:', error);
              }
            }
          }
        }
      }
    };
    
    this.msalClient = new PublicClientApplication(msalConfig);
  }

  async authenticate(): Promise<void> {
    // Try to get cached account
    const accounts = await this.msalClient.getAllAccounts();
    
    if (accounts.length > 0) {
      this.currentAccount = accounts[0];
      console.log(`Using cached credentials for: ${this.currentAccount.username}`);
      
      // Try silent token acquisition
      try {
        await this.acquireTokenSilent();
        console.log('Successfully authenticated using cached credentials');
        return;
      } catch (error) {
        console.log('Silent token acquisition failed, initiating interactive login...');
      }
    }

    // Interactive login required
    await this.interactiveLogin();
  }

  private async interactiveLogin(): Promise<void> {
    const app = express();
    const port = 3000;
    
    return new Promise((resolve, reject) => {
      app.get('/redirect', async (req, res) => {
        const code = req.query.code as string;
        
        if (!code) {
          res.send('Authentication failed - no code received');
          reject(new Error('No authorization code received'));
          return;
        }

        try {
          const tokenResponse = await this.msalClient.acquireTokenByCode({
            code,
            scopes: [
              'User.Read',
              'Calendars.ReadWrite',
              'Mail.Send',
              'Chat.ReadWrite',
              'ChannelMessage.Send',
              'Team.ReadBasic.All',
              'OnlineMeetings.ReadWrite'
            ],
            redirectUri: 'http://localhost:3000/redirect'
          });

          this.currentAccount = tokenResponse.account;
          
          // Note: MSAL-node handles refresh tokens internally for security
          // We rely on the token cache for persistence
          console.log('Authentication tokens cached successfully');

          res.send(`
            <html>
              <body>
                <h1>Authentication successful!</h1>
                <p>Welcome ${tokenResponse.account?.username}</p>
                <p>You can close this window and return to your MCP client.</p>
                <script>window.close();</script>
              </body>
            </html>
          `);

          server.close();
          console.log('Interactive authentication completed successfully');
          resolve();
        } catch (error) {
          console.error('Authentication error:', error);
          res.send(`
            <html>
              <body>
                <h1>Authentication failed</h1>
                <p>Error: ${error instanceof Error ? error.message : 'Unknown error'}</p>
              </body>
            </html>
          `);
          reject(error);
        }
      });

      const server = app.listen(port, async () => {
        try {
          const authUrl = await this.msalClient.getAuthCodeUrl({
            scopes: [
              'User.Read',
              'Calendars.ReadWrite',
              'Mail.Send',
              'Chat.ReadWrite',
              'ChannelMessage.Send',
              'Team.ReadBasic.All',
              'OnlineMeetings.ReadWrite'
            ],
            redirectUri: 'http://localhost:3000/redirect',
            prompt: 'select_account'
          });

          console.log(`Opening browser for authentication...`);
          console.log(`If browser doesn't open automatically, visit: ${authUrl}`);
          await open(authUrl);
        } catch (error) {
          console.error('Failed to initiate authentication:', error);
          reject(error);
        }
      });

      // Timeout after 5 minutes
      setTimeout(() => {
        server.close();
        reject(new Error('Authentication timeout - no response received within 5 minutes'));
      }, 5 * 60 * 1000);
    });
  }

  private async acquireTokenSilent(): Promise<AuthenticationResult> {
    if (!this.currentAccount) {
      throw new Error('No account available for silent token acquisition');
    }
    
    const silentRequest = {
      account: this.currentAccount,
      scopes: [
        'User.Read',
        'Calendars.ReadWrite',
        'Mail.Send',
        'Chat.ReadWrite',
        'ChannelMessage.Send',
        'Team.ReadBasic.All',
        'OnlineMeetings.ReadWrite'
      ],
      forceRefresh: false
    };

    return await this.msalClient.acquireTokenSilent(silentRequest);
  }

  async getGraphClient(): Promise<Client> {
    if (!this.graphClient) {
      this.graphClient = Client.init({
        authProvider: async (done) => {
          try {
            const tokenResponse = await this.acquireTokenSilent();
            done(null, tokenResponse.accessToken);
          } catch (error) {
            console.error('Failed to acquire access token:', error);
            done(error as Error, null);
          }
        }
      });
    }
    
    return this.graphClient;
  }

  async validateTenant(): Promise<boolean> {
    // Validate against Cypherdyne tenant from darbot_config
    if (!this.currentAccount) {
      console.error('No account available for tenant validation');
      return false;
    }
    
    const tenantId = this.currentAccount.tenantId;
    const expectedTenantId = '6b104499-c49f-45dc-b3a2-df95efd6eeb4';
    
    if (tenantId !== expectedTenantId) {
      console.error(`Invalid tenant. Expected: ${expectedTenantId}, Got: ${tenantId}`);
      return false;
    }
    
    console.log(`Tenant validation successful: ${tenantId}`);
    return true;
  }

  async validateUser(): Promise<boolean> {
    // Validate that the user is from the authorized domain
    if (!this.currentAccount) {
      console.error('No account available for user validation');
      return false;
    }
    
    const username = this.currentAccount.username;
    if (!username.endsWith('@timelarp.com')) {
      console.error(`Invalid user domain. Expected @timelarp.com, Got: ${username}`);
      return false;
    }
    
    console.log(`User validation successful: ${username}`);
    return true;
  }

  getCurrentUser(): AccountInfo | null {
    return this.currentAccount;
  }

  async signOut(): Promise<void> {
    if (this.currentAccount) {
      try {
        // Remove refresh token from keychain
        await keytar.deletePassword(SERVICE_NAME, this.currentAccount.username);
      } catch (error) {
        console.warn('Failed to remove refresh token from keychain:', error);
      }
      
      // Clear token cache
      try {
        await fs.unlink(TOKEN_CACHE_FILE);
      } catch (error) {
        console.warn('Failed to remove token cache file:', error);
      }
      
      this.currentAccount = null;
      this.graphClient = null;
      console.log('Sign out completed');
    }
  }
}