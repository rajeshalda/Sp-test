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
export declare class GroupMembershipService {
    private graphClient;
    private _userSite;
    private _syncStatus;
    constructor(graphClient: MSGraphClientV3);
    getUserGroupMemberships(): Promise<IGroupMembership[]>;
    getConnectedTeams(): Promise<IConnectedTeam[]>;
    getUserFilesInTeams(): Promise<IUserFile[]>;
    getUserPersonalSite(): Promise<IUserSite>;
    syncFilesToUserSite(): Promise<void>;
    getSyncStatus(): ISyncStatus;
    toggleSync(enabled: boolean): Promise<void>;
    startBackgroundSync(): Promise<void>;
    private _getAllFilesFromDrive;
    private _copyFileToUserSite;
    private _groupFilesByTeam;
    private _getFileType;
    private _getSyncPreference;
    private _saveSyncPreference;
    private _saveSyncStatus;
    private _loadSyncStatus;
    private _handleError;
}
//# sourceMappingURL=GroupMembershipService.d.ts.map