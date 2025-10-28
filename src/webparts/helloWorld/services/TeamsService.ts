import { MSGraphClientV3 } from '@microsoft/sp-http';

export interface ITeam {
  id: string;
  displayName: string;
  description?: string;
}

export interface IChannel {
  id: string;
  displayName: string;
  description?: string;
  teamId: string;
  teamName: string;
}

export interface ITreeNode {
  id: string;
  name: string;
  type: 'team' | 'channel' | 'folder' | 'file';
  description?: string;
  parentId?: string;
  teamId?: string;
  channelId?: string;
  webUrl?: string;
  size?: number;
  children?: ITreeNode[];
  isExpanded?: boolean;
  isLoading?: boolean;
  childCount?: number;
  driveId?: string;
  itemId?: string;
}

export interface IDriveItem {
  id: string;
  name: string;
  webUrl?: string;
  size?: number;
  folder?: {
    childCount: number;
  };
  file?: {
    mimeType: string;
  };
  createdDateTime?: string;
  lastModifiedDateTime?: string;
}

export interface ITeamsServiceError {
  type: 'NO_PERMISSIONS' | 'NO_TEAMS' | 'NETWORK_ERROR' | 'UNKNOWN';
  message: string;
  originalError?: any;
}

export class TeamsService {
  private graphClient: MSGraphClientV3;

  constructor(graphClient: MSGraphClientV3) {
    this.graphClient = graphClient;
  }

  public async getUserTeams(): Promise<ITeam[]> {
    try {
      const response = await this.graphClient
        .api('/me/joinedTeams')
        .get();

      if (!response || !response.value) {
        throw new Error('No teams data received');
      }

      return response.value.map((team: any) => ({
        id: team.id,
        displayName: team.displayName,
        description: team.description
      }));
    } catch (error) {
      throw this._handleError(error);
    }
  }

  public async getTeamChannels(teamId: string): Promise<Omit<IChannel, 'teamId' | 'teamName'>[]> {
    try {
      const response = await this.graphClient
        .api(`/teams/${teamId}/channels`)
        .get();

      if (!response || !response.value) {
        throw new Error('No channels data received');
      }

      return response.value.map((channel: any) => ({
        id: channel.id,
        displayName: channel.displayName,
        description: channel.description
      }));
    } catch (error) {
      throw this._handleError(error);
    }
  }

  public async getAllChannelsForUser(): Promise<IChannel[]> {
    try {
      const teams = await this.getUserTeams();

      if (teams.length === 0) {
        const error: ITeamsServiceError = {
          type: 'NO_TEAMS',
          message: 'User is not a member of any teams'
        };
        throw error;
      }

      const allChannels: IChannel[] = [];

      for (const team of teams) {
        try {
          const channels = await this.getTeamChannels(team.id);
          const channelsWithTeamInfo = channels.map(channel => ({
            ...channel,
            teamId: team.id,
            teamName: team.displayName
          }));
          allChannels.push(...channelsWithTeamInfo);
        } catch (error) {
          console.warn(`Failed to get channels for team ${team.displayName}:`, error);
        }
      }

      return allChannels;
    } catch (error) {
      if ((error as ITeamsServiceError).type) {
        throw error;
      }
      throw this._handleError(error);
    }
  }

  public async getChannelFilesFolder(teamId: string, channelId: string): Promise<IDriveItem> {
    try {
      const response = await this.graphClient
        .api(`/teams/${teamId}/channels/${channelId}/filesFolder`)
        .get();

      console.log('Raw filesFolder response:', JSON.stringify(response, null, 2));

      return {
        id: response.id,
        name: response.name,
        webUrl: response.webUrl,
        size: response.size,
        folder: response.folder,
        createdDateTime: response.createdDateTime,
        lastModifiedDateTime: response.lastModifiedDateTime
      };
    } catch (error) {
      throw this._handleError(error);
    }
  }

  public async getDriveItemChildren(driveId: string, itemId: string = 'root'): Promise<IDriveItem[]> {
    try {
      const response = await this.graphClient
        .api(`/drives/${driveId}/items/${itemId}/children`)
        .get();

      if (!response || !response.value) {
        return [];
      }

      return response.value.map((item: any) => ({
        id: item.id,
        name: item.name,
        webUrl: item.webUrl,
        size: item.size,
        folder: item.folder,
        file: item.file,
        createdDateTime: item.createdDateTime,
        lastModifiedDateTime: item.lastModifiedDateTime
      }));
    } catch (error) {
      console.warn('Failed to get drive item children:', error);
      return [];
    }
  }

  public async buildTeamsTreeView(): Promise<ITreeNode[]> {
    try {
      const teams = await this.getUserTeams();
      const treeNodes: ITreeNode[] = [];

      for (const team of teams) {
        const teamNode: ITreeNode = {
          id: team.id,
          name: team.displayName,
          type: 'team',
          description: team.description,
          isExpanded: false,
          children: []
        };

        try {
          const channels = await this.getTeamChannels(team.id);

          for (const channel of channels) {
            const channelNode: ITreeNode = {
              id: `${team.id}_${channel.id}`,
              name: channel.displayName,
              type: 'channel',
              description: channel.description,
              parentId: team.id,
              teamId: team.id,
              channelId: channel.id,
              isExpanded: false,
              children: [],
              childCount: 0,
              isLoading: false
            };

            if (teamNode.children) {
              teamNode.children.push(channelNode);
            }
          }
        } catch (error) {
          console.warn(`Failed to get channels for team ${team.displayName}:`, error);
        }

        treeNodes.push(teamNode);
      }

      return treeNodes;
    } catch (error) {
      throw this._handleError(error);
    }
  }

  public async loadChannelFiles(teamId: string, channelId: string): Promise<ITreeNode[]> {
    try {
      // First, get the channel's filesFolder which gives us the SharePoint location
      const filesFolder = await this.getChannelFilesFolder(teamId, channelId);

      console.log('FilesFolder response:', filesFolder);

      // Extract the drive and item information from the filesFolder
      // The filesFolder should contain drive information
      if (!filesFolder.webUrl) {
        console.warn('No webUrl in filesFolder response');
        return [];
      }

      // Try direct approach using the team's drive and look for files
      // Method 1: Use the team group drive approach
      const driveResponse = await this.graphClient
        .api(`/groups/${teamId}/drive`)
        .get();

      if (!driveResponse) {
        console.warn('No drive response for team');
        return [];
      }

      console.log('Drive response:', driveResponse);

      // Method 2: Get all items in the drive root and look for channel folders
      const rootItems = await this.getDriveItemChildren(driveResponse.id, 'root');
      console.log('Root items:', rootItems);

      // Look for a folder that matches the channel name or try General folder first
      let channelFolderItems: IDriveItem[] = [];

      // Try to find files directly in the main Documents folder or subfolders
      for (const item of rootItems) {
        if (item.folder) {
          const folderContents = await this.getDriveItemChildren(driveResponse.id, item.id);
          console.log(`Contents of ${item.name}:`, folderContents);
          channelFolderItems = channelFolderItems.concat(folderContents);
        }
      }

      if (channelFolderItems.length === 0) {
        // Fallback: try to get items directly from filesFolder if it has a drive reference
        try {
          // Extract drive ID and item ID from the response if available
          const folderData = filesFolder as any;
          if (folderData.parentReference && folderData.parentReference.driveId) {
            channelFolderItems = await this.getDriveItemChildren(
              folderData.parentReference.driveId,
              folderData.id
            );
          }
        } catch (folderError) {
          console.warn('Fallback method failed:', folderError);
        }
      }

      return this._convertDriveItemsToTreeNodes(channelFolderItems, driveResponse.id, `${teamId}_${channelId}`);
    } catch (error) {
      console.error(`Failed to load files for channel ${channelId}:`, error);
      return [];
    }
  }

  public async loadChannelFilesDirect(teamId: string, channelId: string): Promise<ITreeNode[]> {
    try {
      console.log(`🔍 Starting loadChannelFilesDirect for teamId: ${teamId}, channelId: ${channelId}`);

      // Step 1: Get the filesFolder first
      const filesFolder = await this.getChannelFilesFolder(teamId, channelId);
      console.log('📁 FilesFolder data:', JSON.stringify(filesFolder, null, 2));

      // Step 2: Try to get children using drive information from filesFolder
      if (filesFolder.id) {
        // Extract drive info from the response
        const fullResponse = await this.graphClient
          .api(`/teams/${teamId}/channels/${channelId}/filesFolder`)
          .select('id,name,parentReference,webUrl')
          .get();

        console.log('🔧 Full filesFolder response:', JSON.stringify(fullResponse, null, 2));

        if (fullResponse.parentReference && fullResponse.parentReference.driveId) {
          console.log(`📂 Found drive ID: ${fullResponse.parentReference.driveId}, item ID: ${fullResponse.id}`);

          // Get the children of this folder
          const children = await this.getDriveItemChildren(fullResponse.parentReference.driveId, fullResponse.id);
          console.log(`📄 Found ${children.length} items in channel folder`);

          if (children.length > 0) {
            console.log('✅ Successfully loaded files:', children.map(c => c.name));
            return this._convertDriveItemsToTreeNodes(children, fullResponse.parentReference.driveId, `${teamId}_${channelId}`);
          }
        }
      }

      console.log('⚠️ No files found or no drive reference available');
      return [];
    } catch (error) {
      console.error(`❌ Direct approach failed for channel ${channelId}:`, error);
      console.error('Error details:', JSON.stringify(error, null, 2));
      return [];
    }
  }

  private _convertDriveItemsToTreeNodes(driveItems: IDriveItem[], driveId: string, parentId: string): ITreeNode[] {
    return driveItems.map(item => ({
      id: `${parentId}_${item.id}`,
      name: item.name,
      type: item.folder ? 'folder' : 'file',
      parentId: parentId,
      webUrl: item.webUrl,
      size: item.size,
      driveId: driveId,
      itemId: item.id,
      childCount: item.folder?.childCount || 0,
      isExpanded: false,
      children: [],
      isLoading: false
    }));
  }

  private _handleError(error: any): ITeamsServiceError {
    console.error('Teams Service Error:', error);

    if (error.code === 'Forbidden' || error.status === 403) {
      return {
        type: 'NO_PERMISSIONS',
        message: 'Insufficient permissions to access Teams data. Please contact your administrator.',
        originalError: error
      };
    }

    if (error.code === 'NetworkError' || error.name === 'NetworkError') {
      return {
        type: 'NETWORK_ERROR',
        message: 'Network error occurred while fetching Teams data. Please check your connection and try again.',
        originalError: error
      };
    }

    return {
      type: 'UNKNOWN',
      message: `An unexpected error occurred: ${error.message || 'Unknown error'}`,
      originalError: error
    };
  }
}