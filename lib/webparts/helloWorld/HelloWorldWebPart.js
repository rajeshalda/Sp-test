var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        if (typeof b !== "function" && b !== null)
            throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (g && (g = 0, op[0] && (_ = 0)), _) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
import { Version } from '@microsoft/sp-core-library';
import { PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import styles from './HelloWorldWebPart.module.scss';
import * as strings from 'HelloWorldWebPartStrings';
import { GroupMembershipService } from './services/GroupMembershipService';
var HelloWorldWebPart = /** @class */ (function (_super) {
    __extends(HelloWorldWebPart, _super);
    function HelloWorldWebPart() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this._isDarkTheme = false;
        _this._environmentMessage = '';
        _this._clickCount = 0;
        _this._groupService = null;
        _this._teams = [];
        _this._userFiles = [];
        _this._syncStatus = null;
        _this._isLoading = false;
        _this._error = '';
        _this._userSiteUrl = '';
        _this._backgroundSyncTimer = null;
        return _this;
    }
    HelloWorldWebPart.prototype.render = function () {
        this.domElement.innerHTML = "\n    <section class=\"".concat(styles.helloWorld, " ").concat(!!this.context.sdks.microsoftTeams ? styles.teams : '', "\">\n      <div class=\"").concat(styles.headerCard, "\">\n        <div class=\"").concat(styles.welcome, "\">\n          <img alt=\"\" src=\"").concat(this._isDarkTheme ? require('./assets/welcome-dark.png') : require('./assets/welcome-light.png'), "\" class=\"").concat(styles.welcomeImage, "\" />\n          <h2>Well done, ").concat(escape(this.context.pageContext.user.displayName), "!</h2>\n          <div class=\"").concat(styles.environmentInfo, "\">").concat(this._environmentMessage, "</div>\n          <div class=\"").concat(styles.propertyInfo, "\">Web part property: <strong>").concat(escape(this.properties.description), "</strong></div>\n        </div>\n      </div>\n\n      <div class=\"").concat(styles.teamsCard, "\">\n        <h3>\uD83D\uDCC1 Teams File Sync Manager</h3>\n        <div class=\"").concat(styles.teamsSection, "\">\n          ").concat(this._renderSyncInterface(), "\n        </div>\n      </div>\n\n      <div class=\"").concat(styles.interactiveCard, "\">\n        <h3>\uD83C\uDFAF Interactive Test Area</h3>\n        <div class=\"").concat(styles.counterSection, "\">\n          <p>Click counter: <span class=\"").concat(styles.counterDisplay, "\">").concat(this._clickCount, "</span></p>\n          <button class=\"").concat(styles.primaryButton, "\" data-action=\"increment\">Increment Counter</button>\n          <button class=\"").concat(styles.secondaryButton, "\" data-action=\"reset\">Reset Counter</button>\n        </div>\n      </div>\n\n      <div class=\"").concat(styles.contentCard, "\">\n        <h3>\uD83D\uDCDA Welcome to SharePoint Framework!</h3>\n        <p class=\"").concat(styles.description, "\">\n        The SharePoint Framework (SPFx) is an extensibility model for Microsoft Viva, Microsoft Teams and SharePoint. It's the easiest way to extend Microsoft 365 with automatic Single Sign On, automatic hosting and industry standard tooling.\n        </p>\n        <h4>\uD83D\uDE80 Learn more about SPFx development:</h4>\n          <ul class=\"").concat(styles.links, "\">\n            <li><a href=\"https://aka.ms/spfx\" target=\"_blank\">\uD83D\uDCD6 SharePoint Framework Overview</a></li>\n            <li><a href=\"https://aka.ms/spfx-yeoman-graph\" target=\"_blank\">\uD83D\uDCCA Use Microsoft Graph in your solution</a></li>\n            <li><a href=\"https://aka.ms/spfx-yeoman-teams\" target=\"_blank\">\uD83D\uDC65 Build for Microsoft Teams using SharePoint Framework</a></li>\n            <li><a href=\"https://aka.ms/spfx-yeoman-viva\" target=\"_blank\">\uD83D\uDCBC Build for Microsoft Viva Connections using SharePoint Framework</a></li>\n            <li><a href=\"https://aka.ms/spfx-yeoman-store\" target=\"_blank\">\uD83C\uDFEA Publish SharePoint Framework applications to the marketplace</a></li>\n            <li><a href=\"https://aka.ms/spfx-yeoman-api\" target=\"_blank\">\uD83D\uDD27 SharePoint Framework API reference</a></li>\n            <li><a href=\"https://aka.ms/m365pnp\" target=\"_blank\">\uFFFD\uD83E\uDD1D Microsoft 365 Developer Community</a></li>\n          </ul>\n      </div>\n    </section>");
        this._bindEvents();
    };
    HelloWorldWebPart.prototype._renderSyncInterface = function () {
        if (this._isLoading) {
            return "\n        <div class=\"".concat(styles.loadingState, "\">\n          <div class=\"").concat(styles.spinner, "\"></div>\n          <p>Loading sync information...</p>\n        </div>\n      ");
        }
        if (this._error) {
            return "\n        <div class=\"".concat(styles.errorState, "\">\n          <p class=\"").concat(styles.errorMessage, "\">\u274C ").concat(this._error, "</p>\n          <button class=\"").concat(styles.primaryButton, "\" data-action=\"retry\">Try Again</button>\n        </div>\n      ");
        }
        if (!this._syncStatus) {
            return "\n        <div class=\"".concat(styles.emptyState, "\">\n          <button class=\"").concat(styles.primaryButton, "\" data-action=\"initialize-sync\">Initialize File Sync</button>\n        </div>\n      ");
        }
        var statusIcon = this._syncStatus.status === 'syncing' ? '🔄' :
            this._syncStatus.status === 'error' ? '❌' :
                this._syncStatus.isEnabled ? '●' : '○';
        var statusText = this._syncStatus.isEnabled ? 'Enabled' : 'Disabled';
        var lastSyncText = this._syncStatus.lastSyncDate ?
            this._syncStatus.lastSyncDate.toLocaleString() : 'Never';
        return "\n      <div class=\"".concat(styles.syncInterface, "\">\n        <div class=\"").concat(styles.syncHeader, "\">\n          <div class=\"").concat(styles.siteInfo, "\">\n            <p><strong>Your SharePoint Site:</strong>\n              ").concat(this._userSiteUrl ? "<a href=\"".concat(this._userSiteUrl, "\" target=\"_blank\">").concat(this._userSiteUrl, "</a>") : 'Loading...', "\n            </p>\n          </div>\n\n          <div class=\"").concat(styles.syncStatus, "\">\n            <p><strong>Sync Status:</strong> ").concat(statusIcon, " ").concat(statusText, "</p>\n            <p><strong>Last Sync:</strong> ").concat(lastSyncText, "</p>\n            <p><strong>Files Synced:</strong> ").concat(this._syncStatus.fileCount, " files from ").concat(this._syncStatus.teamsCount, " teams</p>\n            ").concat(this._syncStatus.error ? "<p class=\"".concat(styles.errorMessage, "\">Error: ").concat(this._syncStatus.error, "</p>") : '', "\n          </div>\n        </div>\n\n        <div class=\"").concat(styles.syncControls, "\">\n          <button class=\"").concat(styles.primaryButton, "\"\n                  data-action=\"").concat(this._syncStatus.isEnabled ? 'disable-sync' : 'enable-sync', "\"\n                  ").concat(this._syncStatus.status === 'syncing' ? 'disabled' : '', ">\n            \uD83D\uDD04 ").concat(this._syncStatus.isEnabled ? 'Disable Sync' : 'Enable Sync', "\n          </button>\n          ").concat(this._syncStatus.isEnabled ? "<button class=\"".concat(styles.secondaryButton, "\" data-action=\"view-files\">\uD83D\uDCC1 View Files</button>") : '', "\n        </div>\n\n        ").concat(this._teams.length > 0 ? this._renderTeamsList() : '', "\n      </div>\n    ");
    };
    HelloWorldWebPart.prototype._renderTeamsList = function () {
        if (this._teams.length === 0)
            return '';
        var html = "\n      <div class=\"".concat(styles.teamsListContainer, "\">\n        <h4>Teams Being Synced:</h4>\n        <div class=\"").concat(styles.teamsList, "\">\n    ");
        var _loop_1 = function (team) {
            var teamFiles = this_1._userFiles.filter(function (f) { return f.teamId === team.id; });
            var fileCount = teamFiles.length;
            html += "\n        <div class=\"".concat(styles.teamItem, "\">\n          <span class=\"").concat(styles.teamIcon, "\">\u2705</span>\n          <span class=\"").concat(styles.teamName, "\">").concat(escape(team.displayName), "</span>\n          <span class=\"").concat(styles.fileCount, "\">(").concat(fileCount, " files where you contributed)</span>\n        </div>\n      ");
        };
        var this_1 = this;
        for (var _i = 0, _a = this._teams; _i < _a.length; _i++) {
            var team = _a[_i];
            _loop_1(team);
        }
        html += "\n        </div>\n      </div>\n    ";
        return html;
    };
    HelloWorldWebPart.prototype._bindEvents = function () {
        var _this = this;
        this.domElement.addEventListener('click', function (event) {
            var target = event.target;
            var action = target.getAttribute('data-action');
            if (action === 'increment') {
                _this._clickCount++;
                _this._updateCounter();
            }
            else if (action === 'reset') {
                _this._clickCount = 0;
                _this._updateCounter();
            }
            else if (action === 'initialize-sync') {
                _this._initializeSync().catch(console.error);
            }
            else if (action === 'enable-sync') {
                _this._toggleSync(true).catch(console.error);
            }
            else if (action === 'disable-sync') {
                _this._toggleSync(false).catch(console.error);
            }
            else if (action === 'view-files') {
                _this._viewSyncedFiles();
            }
            else if (action === 'retry') {
                _this._error = '';
                _this._initializeSync().catch(console.error);
            }
        });
    };
    HelloWorldWebPart.prototype._updateCounter = function () {
        var counterDisplay = this.domElement.querySelector(".".concat(styles.counterDisplay));
        if (counterDisplay) {
            counterDisplay.textContent = this._clickCount.toString();
        }
    };
    HelloWorldWebPart.prototype._initializeSync = function () {
        return __awaiter(this, void 0, void 0, function () {
            var graphClient, userSite, _a, _b, error_1, syncError;
            return __generator(this, function (_c) {
                switch (_c.label) {
                    case 0:
                        this._isLoading = true;
                        this._error = '';
                        this.render();
                        _c.label = 1;
                    case 1:
                        _c.trys.push([1, 7, , 8]);
                        if (!!this._groupService) return [3 /*break*/, 3];
                        return [4 /*yield*/, this.context.msGraphClientFactory.getClient('3')];
                    case 2:
                        graphClient = _c.sent();
                        this._groupService = new GroupMembershipService(graphClient);
                        _c.label = 3;
                    case 3: return [4 /*yield*/, this._groupService.getUserPersonalSite()];
                    case 4:
                        userSite = _c.sent();
                        this._userSiteUrl = userSite.webUrl;
                        _a = this;
                        return [4 /*yield*/, this._groupService.getConnectedTeams()];
                    case 5:
                        _a._teams = _c.sent();
                        _b = this;
                        return [4 /*yield*/, this._groupService.getUserFilesInTeams()];
                    case 6:
                        _b._userFiles = _c.sent();
                        this._syncStatus = this._groupService.getSyncStatus();
                        this._isLoading = false;
                        this.render();
                        return [3 /*break*/, 8];
                    case 7:
                        error_1 = _c.sent();
                        this._isLoading = false;
                        syncError = error_1;
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
                        console.error('Error initializing sync:', error_1);
                        return [3 /*break*/, 8];
                    case 8: return [2 /*return*/];
                }
            });
        });
    };
    HelloWorldWebPart.prototype._toggleSync = function (enabled) {
        return __awaiter(this, void 0, void 0, function () {
            var error_2;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (!this._groupService)
                            return [2 /*return*/];
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 3, , 4]);
                        this._isLoading = true;
                        this.render();
                        return [4 /*yield*/, this._groupService.toggleSync(enabled)];
                    case 2:
                        _a.sent();
                        this._syncStatus = this._groupService.getSyncStatus();
                        if (enabled) {
                            this._startBackgroundSync();
                        }
                        else {
                            this._stopBackgroundSync();
                        }
                        this._isLoading = false;
                        this.render();
                        return [3 /*break*/, 4];
                    case 3:
                        error_2 = _a.sent();
                        this._isLoading = false;
                        this._error = "Failed to ".concat(enabled ? 'enable' : 'disable', " sync: ").concat(error_2.message);
                        this.render();
                        console.error('Error toggling sync:', error_2);
                        return [3 /*break*/, 4];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    HelloWorldWebPart.prototype._viewSyncedFiles = function () {
        if (this._userSiteUrl) {
            var syncFolderUrl = "".concat(this._userSiteUrl, "/Teams File Sync");
            window.open(syncFolderUrl, '_blank');
        }
    };
    HelloWorldWebPart.prototype._startBackgroundSync = function () {
        var _this = this;
        this._stopBackgroundSync();
        this._backgroundSyncTimer = window.setInterval(function () { return __awaiter(_this, void 0, void 0, function () {
            var error_3;
            var _a;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        if (!(this._groupService && ((_a = this._syncStatus) === null || _a === void 0 ? void 0 : _a.isEnabled))) return [3 /*break*/, 4];
                        _b.label = 1;
                    case 1:
                        _b.trys.push([1, 3, , 4]);
                        console.log('Running background sync...');
                        return [4 /*yield*/, this._groupService.startBackgroundSync()];
                    case 2:
                        _b.sent();
                        this._syncStatus = this._groupService.getSyncStatus();
                        this.render();
                        return [3 /*break*/, 4];
                    case 3:
                        error_3 = _b.sent();
                        console.error('Background sync failed:', error_3);
                        return [3 /*break*/, 4];
                    case 4: return [2 /*return*/];
                }
            });
        }); }, 4 * 60 * 60 * 1000);
    };
    HelloWorldWebPart.prototype._stopBackgroundSync = function () {
        if (this._backgroundSyncTimer) {
            window.clearInterval(this._backgroundSyncTimer);
            this._backgroundSyncTimer = null;
        }
    };
    HelloWorldWebPart.prototype.onInit = function () {
        var _this = this;
        return this._getEnvironmentMessage().then(function (message) {
            _this._environmentMessage = message;
            // Don't auto-initialize on page load to prevent immediate API throttling
            // User must click "Initialize File Sync" button to start
            // Only restore background sync if it was previously enabled
            if (_this.context.msGraphClientFactory) {
                var syncEnabled = _this._getSyncPreferenceFromStorage();
                if (syncEnabled) {
                    // Initialize the service but don't trigger sync immediately
                    _this.context.msGraphClientFactory.getClient('3').then(function (graphClient) {
                        _this._groupService = new GroupMembershipService(graphClient);
                        _this._syncStatus = _this._groupService.getSyncStatus();
                        // Start background sync on a delay to avoid initial load throttling
                        setTimeout(function () {
                            var _a;
                            if ((_a = _this._syncStatus) === null || _a === void 0 ? void 0 : _a.isEnabled) {
                                _this._startBackgroundSync();
                            }
                        }, 5000); // 5 second delay
                    }).catch(console.error);
                }
            }
        });
    };
    HelloWorldWebPart.prototype._getSyncPreferenceFromStorage = function () {
        try {
            var stored = localStorage.getItem('teamsFileSyncEnabled');
            return stored === 'true';
        }
        catch (_a) {
            return false;
        }
    };
    HelloWorldWebPart.prototype._getEnvironmentMessage = function () {
        var _this = this;
        if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
            return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
                .then(function (context) {
                var environmentMessage = '';
                switch (context.app.host.name) {
                    case 'Office': // running in Office
                        environmentMessage = _this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
                        break;
                    case 'Outlook': // running in Outlook
                        environmentMessage = _this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
                        break;
                    case 'Teams': // running in Teams
                    case 'TeamsModern':
                        environmentMessage = _this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
                        break;
                    default:
                        environmentMessage = strings.UnknownEnvironment;
                }
                return environmentMessage;
            });
        }
        return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
    };
    HelloWorldWebPart.prototype.onThemeChanged = function (currentTheme) {
        if (!currentTheme) {
            return;
        }
        this._isDarkTheme = !!currentTheme.isInverted;
        var semanticColors = currentTheme.semanticColors;
        if (semanticColors) {
            this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
            this.domElement.style.setProperty('--link', semanticColors.link || null);
            this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
        }
    };
    Object.defineProperty(HelloWorldWebPart.prototype, "dataVersion", {
        get: function () {
            return Version.parse('1.0');
        },
        enumerable: false,
        configurable: true
    });
    HelloWorldWebPart.prototype.onDispose = function () {
        this._stopBackgroundSync();
        _super.prototype.onDispose.call(this);
    };
    HelloWorldWebPart.prototype.getPropertyPaneConfiguration = function () {
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
    };
    return HelloWorldWebPart;
}(BaseClientSideWebPart));
export default HelloWorldWebPart;
//# sourceMappingURL=HelloWorldWebPart.js.map