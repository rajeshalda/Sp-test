import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './HelloWorldWebPart.module.scss';
import * as strings from 'HelloWorldWebPartStrings';
import { GroupMembershipService, IConnectedTeam, IUserFile, ISyncStatus, IGroupMembershipServiceError } from './services/GroupMembershipService';

export interface IHelloWorldWebPartProps {
  description: string;
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _clickCount: number = 0;
  private _groupService: GroupMembershipService | null = null;
  private _teams: IConnectedTeam[] = [];
  private _userFiles: IUserFile[] = [];
  private _syncStatus: ISyncStatus | null = null;
  private _isLoading: boolean = false;
  private _error: string = '';
  private _userSiteUrl: string = '';
  private _backgroundSyncTimer: number | null = null;

  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.helloWorld} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <div class="${styles.headerCard}">
        <div class="${styles.welcome}">
          <img alt="" src="${this._isDarkTheme ? require('./assets/welcome-dark.png') : require('./assets/welcome-light.png')}" class="${styles.welcomeImage}" />
          <h2>Well done, ${escape(this.context.pageContext.user.displayName)}!</h2>
          <div class="${styles.environmentInfo}">${this._environmentMessage}</div>
          <div class="${styles.propertyInfo}">Web part property: <strong>${escape(this.properties.description)}</strong></div>
        </div>
      </div>

      <div class="${styles.teamsCard}">
        <h3>📁 Teams File Sync Manager</h3>
        <div class="${styles.teamsSection}">
          ${this._renderSyncInterface()}
        </div>
      </div>

      <div class="${styles.interactiveCard}">
        <h3>🎯 Interactive Test Area</h3>
        <div class="${styles.counterSection}">
          <p>Click counter: <span class="${styles.counterDisplay}">${this._clickCount}</span></p>
          <button class="${styles.primaryButton}" data-action="increment">Increment Counter</button>
          <button class="${styles.secondaryButton}" data-action="reset">Reset Counter</button>
        </div>
      </div>

      <div class="${styles.contentCard}">
        <h3>📚 Welcome to SharePoint Framework!</h3>
        <p class="${styles.description}">
        The SharePoint Framework (SPFx) is an extensibility model for Microsoft Viva, Microsoft Teams and SharePoint. It's the easiest way to extend Microsoft 365 with automatic Single Sign On, automatic hosting and industry standard tooling.
        </p>
        <h4>🚀 Learn more about SPFx development:</h4>
          <ul class="${styles.links}">
            <li><a href="https://aka.ms/spfx" target="_blank">📖 SharePoint Framework Overview</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-graph" target="_blank">📊 Use Microsoft Graph in your solution</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-teams" target="_blank">👥 Build for Microsoft Teams using SharePoint Framework</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-viva" target="_blank">💼 Build for Microsoft Viva Connections using SharePoint Framework</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-store" target="_blank">🏪 Publish SharePoint Framework applications to the marketplace</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-api" target="_blank">🔧 SharePoint Framework API reference</a></li>
            <li><a href="https://aka.ms/m365pnp" target="_blank">�🤝 Microsoft 365 Developer Community</a></li>
          </ul>
      </div>
    </section>`;

    this._bindEvents();
  }

  private _renderSyncInterface(): string {
    if (this._isLoading) {
      return `
        <div class="${styles.loadingState}">
          <div class="${styles.spinner}"></div>
          <p>Loading sync information...</p>
        </div>
      `;
    }

    if (this._error) {
      return `
        <div class="${styles.errorState}">
          <p class="${styles.errorMessage}">❌ ${this._error}</p>
          <button class="${styles.primaryButton}" data-action="retry">Try Again</button>
        </div>
      `;
    }

    if (!this._syncStatus) {
      return `
        <div class="${styles.emptyState}">
          <button class="${styles.primaryButton}" data-action="initialize-sync">Initialize File Sync</button>
        </div>
      `;
    }

    const statusIcon = this._syncStatus.status === 'syncing' ? '🔄' :
                      this._syncStatus.status === 'error' ? '❌' :
                      this._syncStatus.isEnabled ? '●' : '○';

    const statusText = this._syncStatus.isEnabled ? 'Enabled' : 'Disabled';
    const lastSyncText = this._syncStatus.lastSyncDate ?
      this._syncStatus.lastSyncDate.toLocaleString() : 'Never';

    return `
      <div class="${styles.syncInterface}">
        <div class="${styles.syncHeader}">
          <div class="${styles.siteInfo}">
            <p><strong>Your SharePoint Site:</strong>
              ${this._userSiteUrl ? `<a href="${this._userSiteUrl}" target="_blank">${this._userSiteUrl}</a>` : 'Loading...'}
            </p>
          </div>

          <div class="${styles.syncStatus}">
            <p><strong>Sync Status:</strong> ${statusIcon} ${statusText}</p>
            <p><strong>Last Sync:</strong> ${lastSyncText}</p>
            <p><strong>Files Synced:</strong> ${this._syncStatus.fileCount} files from ${this._syncStatus.teamsCount} teams</p>
            ${this._syncStatus.error ? `<p class="${styles.errorMessage}">Error: ${this._syncStatus.error}</p>` : ''}
          </div>
        </div>

        <div class="${styles.syncControls}">
          <button class="${styles.primaryButton}"
                  data-action="${this._syncStatus.isEnabled ? 'disable-sync' : 'enable-sync'}"
                  ${this._syncStatus.status === 'syncing' ? 'disabled' : ''}>
            🔄 ${this._syncStatus.isEnabled ? 'Disable Sync' : 'Enable Sync'}
          </button>
          ${this._syncStatus.isEnabled ? `<button class="${styles.secondaryButton}" data-action="view-files">📁 View Files</button>` : ''}
        </div>

        ${this._teams.length > 0 ? this._renderTeamsList() : ''}
      </div>
    `;
  }

  private _renderTeamsList(): string {
    if (this._teams.length === 0) return '';

    let html = `
      <div class="${styles.teamsListContainer}">
        <h4>Teams Being Synced:</h4>
        <div class="${styles.teamsList}">
    `;

    for (const team of this._teams) {
      const teamFiles = this._userFiles.filter(f => f.teamId === team.id);
      const fileCount = teamFiles.length;

      html += `
        <div class="${styles.teamItem}">
          <span class="${styles.teamIcon}">✅</span>
          <span class="${styles.teamName}">${escape(team.displayName)}</span>
          <span class="${styles.fileCount}">(${fileCount} files where you contributed)</span>
        </div>
      `;
    }

    html += `
        </div>
      </div>
    `;

    return html;
  }

  private _bindEvents(): void {
    this.domElement.addEventListener('click', (event: Event) => {
      const target = event.target as HTMLElement;
      const action = target.getAttribute('data-action');

      if (action === 'increment') {
        this._clickCount++;
        this._updateCounter();
      } else if (action === 'reset') {
        this._clickCount = 0;
        this._updateCounter();
      } else if (action === 'initialize-sync') {
        this._initializeSync().catch(console.error);
      } else if (action === 'enable-sync') {
        this._toggleSync(true).catch(console.error);
      } else if (action === 'disable-sync') {
        this._toggleSync(false).catch(console.error);
      } else if (action === 'view-files') {
        this._viewSyncedFiles();
      } else if (action === 'retry') {
        this._error = '';
        this._initializeSync().catch(console.error);
      }
    });
  }

  private _updateCounter(): void {
    const counterDisplay = this.domElement.querySelector(`.${styles.counterDisplay}`);
    if (counterDisplay) {
      counterDisplay.textContent = this._clickCount.toString();
    }
  }

  private async _initializeSync(): Promise<void> {
    this._isLoading = true;
    this._error = '';
    this.render();

    try {
      if (!this._groupService) {
        const graphClient = await this.context.msGraphClientFactory.getClient('3');
        this._groupService = new GroupMembershipService(graphClient);
      }

      const userSite = await this._groupService.getUserPersonalSite();
      this._userSiteUrl = userSite.webUrl;

      this._teams = await this._groupService.getConnectedTeams();
      this._userFiles = await this._groupService.getUserFilesInTeams();
      this._syncStatus = this._groupService.getSyncStatus();

      this._isLoading = false;
      this.render();
    } catch (error) {
      this._isLoading = false;
      const syncError = error as IGroupMembershipServiceError;

      switch (syncError.type) {
        case 'NO_PERMISSIONS':
          this._error = 'Insufficient permissions. Please contact your administrator to approve Microsoft Graph permissions.';
          break;
        case 'NO_GROUPS':
          this._error = 'You are not a member of any teams with SharePoint sites.';
          break;
        case 'NO_SITE':
          this._error = 'Unable to access your personal SharePoint site.';
          break;
        case 'NETWORK_ERROR':
          this._error = 'Network error. Please check your connection and try again.';
          break;
        default:
          this._error = syncError.message || 'An unexpected error occurred while initializing sync.';
      }

      this.render();
      console.error('Error initializing sync:', error);
    }
  }

  private async _toggleSync(enabled: boolean): Promise<void> {
    if (!this._groupService) return;

    try {
      this._isLoading = true;
      this.render();

      await this._groupService.toggleSync(enabled);
      this._syncStatus = this._groupService.getSyncStatus();

      if (enabled) {
        this._startBackgroundSync();
      } else {
        this._stopBackgroundSync();
      }

      this._isLoading = false;
      this.render();
    } catch (error) {
      this._isLoading = false;
      this._error = `Failed to ${enabled ? 'enable' : 'disable'} sync: ${(error as Error).message}`;
      this.render();
      console.error('Error toggling sync:', error);
    }
  }

  private _viewSyncedFiles(): void {
    if (this._userSiteUrl) {
      const syncFolderUrl = `${this._userSiteUrl}/Teams File Sync`;
      window.open(syncFolderUrl, '_blank');
    }
  }

  private _startBackgroundSync(): void {
    this._stopBackgroundSync();

    this._backgroundSyncTimer = window.setInterval(async () => {
      if (this._groupService && this._syncStatus?.isEnabled) {
        try {
          console.log('Running background sync...');
          await this._groupService.startBackgroundSync();
          this._syncStatus = this._groupService.getSyncStatus();
          this.render();
        } catch (error) {
          console.error('Background sync failed:', error);
        }
      }
    }, 4 * 60 * 60 * 1000);
  }

  private _stopBackgroundSync(): void {
    if (this._backgroundSyncTimer) {
      window.clearInterval(this._backgroundSyncTimer);
      this._backgroundSyncTimer = null;
    }
  }


  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;

      if (this.context.msGraphClientFactory) {
        this._initializeSync().then(() => {
          if (this._syncStatus?.isEnabled) {
            this._startBackgroundSync();
          }
        }).catch(console.error);
      }
    });
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected onDispose(): void {
    this._stopBackgroundSync();
    super.onDispose();
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
