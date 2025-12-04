import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
export interface IHelloWorldWebPartProps {
    description: string;
}
export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {
    private _isDarkTheme;
    private _environmentMessage;
    private _clickCount;
    private _groupService;
    private _teams;
    private _userFiles;
    private _syncStatus;
    private _isLoading;
    private _error;
    private _userSiteUrl;
    private _backgroundSyncTimer;
    render(): void;
    private _renderSyncInterface;
    private _renderTeamsList;
    private _bindEvents;
    private _updateCounter;
    private _initializeSync;
    private _toggleSync;
    private _viewSyncedFiles;
    private _startBackgroundSync;
    private _stopBackgroundSync;
    protected onInit(): Promise<void>;
    private _getSyncPreferenceFromStorage;
    private _getEnvironmentMessage;
    protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void;
    protected get dataVersion(): Version;
    protected onDispose(): void;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
}
//# sourceMappingURL=HelloWorldWebPart.d.ts.map