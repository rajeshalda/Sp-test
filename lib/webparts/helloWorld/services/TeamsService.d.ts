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
export declare class TeamsService {
    private graphClient;
    constructor(graphClient: MSGraphClientV3);
    getUserTeams(): Promise<ITeam[]>;
    getTeamChannels(teamId: string): Promise<Omit<IChannel, 'teamId' | 'teamName'>[]>;
    getAllChannelsForUser(): Promise<IChannel[]>;
    getChannelFilesFolder(teamId: string, channelId: string): Promise<IDriveItem>;
    getDriveItemChildren(driveId: string, itemId?: string): Promise<IDriveItem[]>;
    buildTeamsTreeView(): Promise<ITreeNode[]>;
    loadChannelFiles(teamId: string, channelId: string): Promise<ITreeNode[]>;
    loadChannelFilesDirect(teamId: string, channelId: string): Promise<ITreeNode[]>;
    private _convertDriveItemsToTreeNodes;
    private _handleError;
}
//# sourceMappingURL=TeamsService.d.ts.map