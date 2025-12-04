import { MSGraphClientV3 } from '@microsoft/sp-http';

export interface IUserSite {
  id: string;
  displayName: string;
  webUrl: string;
  siteCollection?: {
    hostname: string;
  };
}

export interface IGroupMembership {
  id: string;
  displayName: string;
  description?: string;
  resourceBehaviorOptions?: string[];
  resourceProvisioningOptions?: string[];
  visibility?: string;
  createdDateTime?: string;
}

export interface IConnectedTeam {
  id: string;
  displayName: string;
  description?: string;
  groupId: string;
  isArchived?: boolean;
}

export interface IUserFile {
  id: string;
  name: string;
  webUrl: string;
  size?: number;
  fileType: string;
  lastModifiedDateTime: string;
  lastModifiedBy: {
    user: {
      displayName: string;
    };
  };
  teamId: string;
  teamName: string;
  channelId?: string;
  channelName?: string;
  driveId: string;
  itemPath: string;
}

export interface ISyncStatus {
  isEnabled: boolean;
  lastSyncDate?: Date;
  fileCount: number;
  teamsCount: number;
  status: 'idle' | 'syncing' | 'error';
  error?: string;
}

export interface IGroupMembershipServiceError {
  type: 'NO_PERMISSIONS' | 'NO_GROUPS' | 'NO_SITE' | 'SYNC_ERROR' | 'NETWORK_ERROR' | 'UNKNOWN';
  message: string;
  originalError?: any;
}

export class GroupMembershipService {
  private graphClient: MSGraphClientV3;
  private _userSite: IUserSite | null = null;
  private _syncStatus: ISyncStatus;
  private readonly DELAY_BETWEEN_CALLS = 300; // 300ms delay between API calls
  private readonly MAX_RETRIES = 3;
  private readonly INITIAL_RETRY_DELAY = 1000; // 1 second

  constructor(graphClient: MSGraphClientV3) {
    this.graphClient = graphClient;
    this._syncStatus = {
      isEnabled: this._getSyncPreference(),
      fileCount: 0,
      teamsCount: 0,
      status: 'idle'
    };
    this._loadSyncStatus();
  }

  // Rate limiting utility - delay between API calls
  private async _delay(ms: number): Promise<void> {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  // Retry logic with exponential backoff for 429 errors
  private async _retryWithBackoff<T>(
    operation: () => Promise<T>,
    retries: number = this.MAX_RETRIES
  ): Promise<T> {
    for (let attempt = 0; attempt <= retries; attempt++) {
      try {
        return await operation();
      } catch (error: any) {
        const isThrottled = error.statusCode === 429 ||
                          error.code === 'TooManyRequests' ||
                          (error.message && error.message.includes('throttled'));

        if (isThrottled && attempt < retries) {
          const delay = this.INITIAL_RETRY_DELAY * Math.pow(2, attempt);
          const retryAfter = error.retryAfter ? parseInt(error.retryAfter) * 1000 : delay;
          console.warn(`Request throttled. Retrying after ${retryAfter}ms (attempt ${attempt + 1}/${retries})`);
          await this._delay(retryAfter);
        } else {
          throw error;
        }
      }
    }
    throw new Error('Max retries exceeded');
  }

  public async getUserGroupMemberships(): Promise<IGroupMembership[]> {
    try {
      const response = await this.graphClient
        .api('/me/memberOf')
        .select('id,displayName,description,resourceBehaviorOptions,resourceProvisioningOptions,visibility,createdDateTime')
        .get();

      if (!response || !response.value) {
        throw new Error('No group memberships data received');
      }

      // Filter for Teams groups client-side instead of server-side
      const teamGroups = response.value.filter((group: any) =>
        group.resourceProvisioningOptions &&
        group.resourceProvisioningOptions.includes('Team')
      );

      return teamGroups.map((group: any) => ({
        id: group.id,
        displayName: group.displayName,
        description: group.description,
        resourceBehaviorOptions: group.resourceBehaviorOptions,
        resourceProvisioningOptions: group.resourceProvisioningOptions,
        visibility: group.visibility,
        createdDateTime: group.createdDateTime
      }));
    } catch (error) {
      throw this._handleError(error);
    }
  }

  public async getConnectedTeams(): Promise<IConnectedTeam[]> {
    try {
      // Try to get joined teams directly first (single API call - preferred)
      try {
        const teamsResponse = await this._retryWithBackoff(() =>
          this.graphClient
            .api('/me/joinedTeams')
            .select('id,displayName,description,isArchived')
            .get()
        );

        if (teamsResponse && teamsResponse.value && teamsResponse.value.length > 0) {
          return teamsResponse.value.map((team: any) => ({
            id: team.id,
            displayName: team.displayName,
            description: team.description,
            groupId: team.id, // Team ID is the same as Group ID
            isArchived: team.isArchived || false
          }));
        }
      } catch (directError) {
        console.warn('Direct teams API failed, trying via groups:', directError);
      }

      // Fallback to groups approach (multiple API calls - add delays)
      const groups = await this.getUserGroupMemberships();
      const teams: IConnectedTeam[] = [];

      for (let i = 0; i < groups.length; i++) {
        const group = groups[i];
        try {
          // Add delay before each call to avoid throttling
          if (i > 0) {
            await this._delay(this.DELAY_BETWEEN_CALLS);
          }

          const teamResponse = await this._retryWithBackoff(() =>
            this.graphClient
              .api(`/teams/${group.id}`)
              .select('id,displayName,description,isArchived')
              .get()
          );

          teams.push({
            id: teamResponse.id,
            displayName: teamResponse.displayName,
            description: teamResponse.description,
            groupId: group.id,
            isArchived: teamResponse.isArchived
          });
        } catch (teamError) {
          console.warn(`Group ${group.displayName} is not a team or access denied:`, teamError);
        }
      }

      if (teams.length === 0) {
        const error: IGroupMembershipServiceError = {
          type: 'NO_GROUPS',
          message: 'User is not a member of any teams with SharePoint sites'
        };
        throw error;
      }

      return teams;
    } catch (error) {
      if ((error as IGroupMembershipServiceError).type) {
        throw error;
      }
      throw this._handleError(error);
    }
  }

  public async getUserFilesInTeams(): Promise<IUserFile[]> {
    try {
      const teams = await this.getConnectedTeams();
      const userFiles: IUserFile[] = [];

      for (let i = 0; i < teams.length; i++) {
        const team = teams[i];
        try {
          // Add delay between team operations
          if (i > 0) {
            await this._delay(this.DELAY_BETWEEN_CALLS);
          }

          const driveResponse = await this._retryWithBackoff(() =>
            this.graphClient
              .api(`/groups/${team.groupId}/drive`)
              .get()
          );

          if (!driveResponse) continue;

          // Pass depth limit to prevent excessive recursion
          const allFiles = await this._getAllFilesFromDrive(driveResponse.id, 'root', '', 2);

          // Sync ALL files from channels you have access to (no filtering by creator/modifier)
          // If you're a member of the team/channel, you can access these files
          const mappedFiles: IUserFile[] = allFiles.map(file => ({
            id: file.id,
            name: file.name,
            webUrl: file.webUrl,
            size: file.size,
            fileType: this._getFileType(file.name),
            lastModifiedDateTime: file.lastModifiedDateTime,
            lastModifiedBy: file.lastModifiedBy || file.createdBy,
            teamId: team.id,
            teamName: team.displayName,
            driveId: driveResponse.id,
            itemPath: file.parentReference?.path || '/'
          }));

          userFiles.push(...mappedFiles);
        } catch (error) {
          console.warn(`Failed to get files for team ${team.displayName}:`, error);
        }
      }

      return userFiles;
    } catch (error) {
      throw this._handleError(error);
    }
  }

  public async getUserPersonalSite(): Promise<IUserSite> {
    try {
      if (this._userSite) {
        return this._userSite;
      }

      const response = await this.graphClient
        .api('/me/drive')
        .select('id,name,webUrl')
        .get();

      if (!response) {
        const error: IGroupMembershipServiceError = {
          type: 'NO_SITE',
          message: 'Unable to access user personal SharePoint site'
        };
        throw error;
      }

      this._userSite = {
        id: response.id,
        displayName: response.name || 'My Site',
        webUrl: response.webUrl
      };

      return this._userSite;
    } catch (error) {
      throw this._handleError(error);
    }
  }

  public async syncFilesToUserSite(): Promise<void> {
    try {
      this._syncStatus.status = 'syncing';
      this._syncStatus.error = undefined;

      await this.getUserPersonalSite();
      const userFiles = await this.getUserFilesInTeams();

      if (userFiles.length === 0) {
        this._syncStatus.status = 'idle';
        return;
      }

      await this.graphClient
        .api('/me/drive')
        .get();

      const syncFolderName = 'Teams File Sync';
      let syncFolderId: string;

      try {
        const existingFolder = await this.graphClient
          .api(`/me/drive/root:/${syncFolderName}`)
          .get();
        syncFolderId = existingFolder.id;
      } catch {
        const newFolder = await this.graphClient
          .api('/me/drive/root/children')
          .post({
            name: syncFolderName,
            folder: {},
            '@microsoft.graph.conflictBehavior': 'replace'
          });
        syncFolderId = newFolder.id;
      }

      const teamGroups = this._groupFilesByTeam(userFiles);
      let totalSynced = 0;
      const teamNames = Object.keys(teamGroups);

      for (let i = 0; i < teamNames.length; i++) {
        const teamName = teamNames[i];
        const files = teamGroups[teamName];
        try {
          // Add delay between team operations
          if (i > 0) {
            await this._delay(this.DELAY_BETWEEN_CALLS);
          }

          let teamFolderId: string;

          try {
            const existingTeamFolder = await this._retryWithBackoff(() =>
              this.graphClient
                .api(`/me/drive/items/${syncFolderId}:/${teamName}`)
                .get()
            );
            teamFolderId = existingTeamFolder.id;
          } catch {
            const newTeamFolder = await this._retryWithBackoff(() =>
              this.graphClient
                .api(`/me/drive/items/${syncFolderId}/children`)
                .post({
                  name: teamName,
                  folder: {},
                  '@microsoft.graph.conflictBehavior': 'replace'
                })
            );
            teamFolderId = newTeamFolder.id;
          }

          // Batch process files to avoid too many concurrent operations
          const batchSize = 5;
          for (let j = 0; j < files.length; j += batchSize) {
            const batch = files.slice(j, j + batchSize);

            for (let k = 0; k < batch.length; k++) {
              try {
                // Add delay between file operations
                if (k > 0) {
                  await this._delay(this.DELAY_BETWEEN_CALLS);
                }
                await this._copyFileToUserSite(batch[k], teamFolderId);
                totalSynced++;
              } catch (copyError) {
                console.warn(`Failed to copy file ${batch[k].name}:`, copyError);
              }
            }

            // Add longer delay between batches
            if (j + batchSize < files.length) {
              await this._delay(this.DELAY_BETWEEN_CALLS * 2);
            }
          }
        } catch (teamError) {
          console.warn(`Failed to sync files for team ${teamName}:`, teamError);
        }
      }

      this._syncStatus = {
        isEnabled: true,
        lastSyncDate: new Date(),
        fileCount: totalSynced,
        teamsCount: Object.keys(teamGroups).length,
        status: 'idle'
      };

      this._saveSyncStatus();
    } catch (error) {
      this._syncStatus.status = 'error';
      this._syncStatus.error = 'Sync failed: ' + (error as Error).message;
      throw this._handleError(error, 'SYNC_ERROR');
    }
  }

  public getSyncStatus(): ISyncStatus {
    return { ...this._syncStatus };
  }

  public async toggleSync(enabled: boolean): Promise<void> {
    this._syncStatus.isEnabled = enabled;
    this._saveSyncPreference(enabled);

    if (enabled) {
      await this.syncFilesToUserSite();
    }
  }

  public async startBackgroundSync(): Promise<void> {
    if (!this._syncStatus.isEnabled) return;

    try {
      await this.syncFilesToUserSite();
    } catch (error) {
      console.error('Background sync failed:', error);
    }
  }

  private async _getAllFilesFromDrive(
    driveId: string,
    itemId: string,
    path: string = '',
    maxDepth: number = 2,
    currentDepth: number = 0
  ): Promise<any[]> {
    try {
      // Stop recursion if max depth reached
      if (currentDepth >= maxDepth) {
        console.log(`Max depth ${maxDepth} reached at path: ${path}`);
        return [];
      }

      // Add delay before API call to prevent throttling
      await this._delay(this.DELAY_BETWEEN_CALLS);

      const response = await this._retryWithBackoff(() =>
        this.graphClient
          .api(`/drives/${driveId}/items/${itemId}/children`)
          .select('id,name,size,webUrl,lastModifiedDateTime,lastModifiedBy,createdBy,parentReference,file,folder')
          .top(100) // Limit items per request
          .get()
      );

      if (!response || !response.value) return [];

      const allFiles: any[] = [];

      for (let i = 0; i < response.value.length; i++) {
        const item = response.value[i];

        if (item.file) {
          item.parentReference = { ...item.parentReference, path: path };
          allFiles.push(item);
        } else if (item.folder && currentDepth < maxDepth - 1) {
          // Only recurse if we haven't reached max depth
          const subFiles = await this._getAllFilesFromDrive(
            driveId,
            item.id,
            `${path}/${item.name}`,
            maxDepth,
            currentDepth + 1
          );
          allFiles.push(...subFiles);
        }
      }

      return allFiles;
    } catch (error: any) {
      // Don't fail completely on throttling - just log and return empty
      if (error.statusCode === 429) {
        console.warn(`Throttled while getting files from drive ${driveId}, item ${itemId}. Skipping this folder.`);
      } else {
        console.warn(`Failed to get files from drive ${driveId}, item ${itemId}:`, error);
      }
      return [];
    }
  }

  private async _copyFileToUserSite(file: IUserFile, destinationFolderId: string): Promise<void> {
    try {
      const sourceUrl = `/drives/${file.driveId}/items/${file.id}`;

      const copyResponse = await this.graphClient
        .api(`${sourceUrl}/copy`)
        .post({
          parentReference: {
            driveId: (await this.graphClient.api('/me/drive').get()).id,
            id: destinationFolderId
          },
          name: file.name
        });

      if (copyResponse && copyResponse.id) {
        console.log(`Successfully copied file: ${file.name}`);
      }
    } catch (error) {
      if ((error as any).code !== 'nameAlreadyExists') {
        throw error;
      }
    }
  }

  private _groupFilesByTeam(files: IUserFile[]): { [teamName: string]: IUserFile[] } {
    return files.reduce((groups, file) => {
      if (!groups[file.teamName]) {
        groups[file.teamName] = [];
      }
      groups[file.teamName].push(file);
      return groups;
    }, {} as { [teamName: string]: IUserFile[] });
  }

  private _getFileType(fileName: string): string {
    const extension = fileName.split('.').pop()?.toLowerCase() || '';
    const types: { [key: string]: string } = {
      'doc': 'Word Document',
      'docx': 'Word Document',
      'xls': 'Excel Spreadsheet',
      'xlsx': 'Excel Spreadsheet',
      'ppt': 'PowerPoint Presentation',
      'pptx': 'PowerPoint Presentation',
      'pdf': 'PDF Document',
      'txt': 'Text File',
      'jpg': 'Image',
      'jpeg': 'Image',
      'png': 'Image',
      'gif': 'Image'
    };
    return types[extension] || 'File';
  }

  private _getSyncPreference(): boolean {
    try {
      const stored = localStorage.getItem('teamsFileSyncEnabled');
      return stored === 'true';
    } catch {
      return false;
    }
  }

  private _saveSyncPreference(enabled: boolean): void {
    try {
      localStorage.setItem('teamsFileSyncEnabled', enabled.toString());
    } catch (error) {
      console.warn('Failed to save sync preference:', error);
    }
  }

  private _saveSyncStatus(): void {
    try {
      const statusToSave = {
        lastSyncDate: this._syncStatus.lastSyncDate?.toISOString(),
        fileCount: this._syncStatus.fileCount,
        teamsCount: this._syncStatus.teamsCount
      };
      localStorage.setItem('teamsFileSyncStatus', JSON.stringify(statusToSave));
    } catch (error) {
      console.warn('Failed to save sync status:', error);
    }
  }

  private _loadSyncStatus(): void {
    try {
      const stored = localStorage.getItem('teamsFileSyncStatus');
      if (stored) {
        const parsed = JSON.parse(stored);
        this._syncStatus.lastSyncDate = parsed.lastSyncDate ? new Date(parsed.lastSyncDate) : undefined;
        this._syncStatus.fileCount = parsed.fileCount || 0;
        this._syncStatus.teamsCount = parsed.teamsCount || 0;
      }
    } catch (error) {
      console.warn('Failed to load sync status:', error);
    }
  }

  private _handleError(error: any, defaultType: IGroupMembershipServiceError['type'] = 'UNKNOWN'): IGroupMembershipServiceError {
    console.error('GroupMembership Service Error:', error);

    if (error.code === 'Forbidden' || error.status === 403) {
      return {
        type: 'NO_PERMISSIONS',
        message: 'Insufficient permissions to access SharePoint data. Please contact your administrator.',
        originalError: error
      };
    }

    if (error.code === 'NetworkError' || error.name === 'NetworkError') {
      return {
        type: 'NETWORK_ERROR',
        message: 'Network error occurred while fetching data. Please check your connection and try again.',
        originalError: error
      };
    }

    if ((error as IGroupMembershipServiceError).type) {
      return error as IGroupMembershipServiceError;
    }

    return {
      type: defaultType,
      message: `An error occurred: ${error.message || 'Unknown error'}`,
      originalError: error
    };
  }
}