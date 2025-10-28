var __assign = (this && this.__assign) || function () {
    __assign = Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
                t[p] = s[p];
        }
        return t;
    };
    return __assign.apply(this, arguments);
};
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
var GroupMembershipService = /** @class */ (function () {
    function GroupMembershipService(graphClient) {
        this._userSite = null;
        this.graphClient = graphClient;
        this._syncStatus = {
            isEnabled: this._getSyncPreference(),
            fileCount: 0,
            teamsCount: 0,
            status: 'idle'
        };
        this._loadSyncStatus();
    }
    GroupMembershipService.prototype.getUserGroupMemberships = function () {
        return __awaiter(this, void 0, void 0, function () {
            var response, teamGroups, error_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, this.graphClient
                                .api('/me/memberOf')
                                .select('id,displayName,description,resourceBehaviorOptions,resourceProvisioningOptions,visibility,createdDateTime')
                                .get()];
                    case 1:
                        response = _a.sent();
                        if (!response || !response.value) {
                            throw new Error('No group memberships data received');
                        }
                        teamGroups = response.value.filter(function (group) {
                            return group.resourceProvisioningOptions &&
                                group.resourceProvisioningOptions.includes('Team');
                        });
                        return [2 /*return*/, teamGroups.map(function (group) { return ({
                                id: group.id,
                                displayName: group.displayName,
                                description: group.description,
                                resourceBehaviorOptions: group.resourceBehaviorOptions,
                                resourceProvisioningOptions: group.resourceProvisioningOptions,
                                visibility: group.visibility,
                                createdDateTime: group.createdDateTime
                            }); })];
                    case 2:
                        error_1 = _a.sent();
                        throw this._handleError(error_1);
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    GroupMembershipService.prototype.getConnectedTeams = function () {
        return __awaiter(this, void 0, void 0, function () {
            var teamsResponse, directError_1, groups, teams, _i, groups_1, group, teamResponse, teamError_1, error, error_2;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 12, , 13]);
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, this.graphClient
                                .api('/me/joinedTeams')
                                .select('id,displayName,description,isArchived')
                                .get()];
                    case 2:
                        teamsResponse = _a.sent();
                        if (teamsResponse && teamsResponse.value && teamsResponse.value.length > 0) {
                            return [2 /*return*/, teamsResponse.value.map(function (team) { return ({
                                    id: team.id,
                                    displayName: team.displayName,
                                    description: team.description,
                                    groupId: team.id, // Team ID is the same as Group ID
                                    isArchived: team.isArchived || false
                                }); })];
                        }
                        return [3 /*break*/, 4];
                    case 3:
                        directError_1 = _a.sent();
                        console.warn('Direct teams API failed, trying via groups:', directError_1);
                        return [3 /*break*/, 4];
                    case 4: return [4 /*yield*/, this.getUserGroupMemberships()];
                    case 5:
                        groups = _a.sent();
                        teams = [];
                        _i = 0, groups_1 = groups;
                        _a.label = 6;
                    case 6:
                        if (!(_i < groups_1.length)) return [3 /*break*/, 11];
                        group = groups_1[_i];
                        _a.label = 7;
                    case 7:
                        _a.trys.push([7, 9, , 10]);
                        return [4 /*yield*/, this.graphClient
                                .api("/teams/".concat(group.id))
                                .select('id,displayName,description,isArchived')
                                .get()];
                    case 8:
                        teamResponse = _a.sent();
                        teams.push({
                            id: teamResponse.id,
                            displayName: teamResponse.displayName,
                            description: teamResponse.description,
                            groupId: group.id,
                            isArchived: teamResponse.isArchived
                        });
                        return [3 /*break*/, 10];
                    case 9:
                        teamError_1 = _a.sent();
                        console.warn("Group ".concat(group.displayName, " is not a team or access denied:"), teamError_1);
                        return [3 /*break*/, 10];
                    case 10:
                        _i++;
                        return [3 /*break*/, 6];
                    case 11:
                        if (teams.length === 0) {
                            error = {
                                type: 'NO_GROUPS',
                                message: 'User is not a member of any teams with SharePoint sites'
                            };
                            throw error;
                        }
                        return [2 /*return*/, teams];
                    case 12:
                        error_2 = _a.sent();
                        if (error_2.type) {
                            throw error_2;
                        }
                        throw this._handleError(error_2);
                    case 13: return [2 /*return*/];
                }
            });
        });
    };
    GroupMembershipService.prototype.getUserFilesInTeams = function () {
        return __awaiter(this, void 0, void 0, function () {
            var teams, userFiles, currentUser_1, _loop_1, this_1, _i, teams_1, team, error_3;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 7, , 8]);
                        return [4 /*yield*/, this.getConnectedTeams()];
                    case 1:
                        teams = _a.sent();
                        userFiles = [];
                        return [4 /*yield*/, this._getCurrentUser()];
                    case 2:
                        currentUser_1 = _a.sent();
                        _loop_1 = function (team) {
                            var driveResponse_1, allFiles, userModifiedFiles, mappedFiles, error_4;
                            return __generator(this, function (_b) {
                                switch (_b.label) {
                                    case 0:
                                        _b.trys.push([0, 3, , 4]);
                                        return [4 /*yield*/, this_1.graphClient
                                                .api("/groups/".concat(team.groupId, "/drive"))
                                                .get()];
                                    case 1:
                                        driveResponse_1 = _b.sent();
                                        if (!driveResponse_1)
                                            return [2 /*return*/, "continue"];
                                        return [4 /*yield*/, this_1._getAllFilesFromDrive(driveResponse_1.id, 'root')];
                                    case 2:
                                        allFiles = _b.sent();
                                        userModifiedFiles = allFiles.filter(function (file) {
                                            var _a, _b, _c, _d, _e, _f, _g, _h, _j, _k, _l, _m;
                                            var isModifiedByUser = ((_b = (_a = file.lastModifiedBy) === null || _a === void 0 ? void 0 : _a.user) === null || _b === void 0 ? void 0 : _b.displayName) === currentUser_1.displayName ||
                                                ((_d = (_c = file.lastModifiedBy) === null || _c === void 0 ? void 0 : _c.user) === null || _d === void 0 ? void 0 : _d.email) === currentUser_1.mail ||
                                                ((_f = (_e = file.lastModifiedBy) === null || _e === void 0 ? void 0 : _e.user) === null || _f === void 0 ? void 0 : _f.id) === currentUser_1.id;
                                            var isCreatedByUser = ((_h = (_g = file.createdBy) === null || _g === void 0 ? void 0 : _g.user) === null || _h === void 0 ? void 0 : _h.displayName) === currentUser_1.displayName ||
                                                ((_k = (_j = file.createdBy) === null || _j === void 0 ? void 0 : _j.user) === null || _k === void 0 ? void 0 : _k.email) === currentUser_1.mail ||
                                                ((_m = (_l = file.createdBy) === null || _l === void 0 ? void 0 : _l.user) === null || _m === void 0 ? void 0 : _m.id) === currentUser_1.id;
                                            return isModifiedByUser || isCreatedByUser;
                                        });
                                        mappedFiles = userModifiedFiles.map(function (file) {
                                            var _a;
                                            return ({
                                                id: file.id,
                                                name: file.name,
                                                webUrl: file.webUrl,
                                                size: file.size,
                                                fileType: _this._getFileType(file.name),
                                                lastModifiedDateTime: file.lastModifiedDateTime,
                                                lastModifiedBy: file.lastModifiedBy || file.createdBy,
                                                teamId: team.id,
                                                teamName: team.displayName,
                                                driveId: driveResponse_1.id,
                                                itemPath: ((_a = file.parentReference) === null || _a === void 0 ? void 0 : _a.path) || '/'
                                            });
                                        });
                                        userFiles.push.apply(userFiles, mappedFiles);
                                        return [3 /*break*/, 4];
                                    case 3:
                                        error_4 = _b.sent();
                                        console.warn("Failed to get files for team ".concat(team.displayName, ":"), error_4);
                                        return [3 /*break*/, 4];
                                    case 4: return [2 /*return*/];
                                }
                            });
                        };
                        this_1 = this;
                        _i = 0, teams_1 = teams;
                        _a.label = 3;
                    case 3:
                        if (!(_i < teams_1.length)) return [3 /*break*/, 6];
                        team = teams_1[_i];
                        return [5 /*yield**/, _loop_1(team)];
                    case 4:
                        _a.sent();
                        _a.label = 5;
                    case 5:
                        _i++;
                        return [3 /*break*/, 3];
                    case 6: return [2 /*return*/, userFiles];
                    case 7:
                        error_3 = _a.sent();
                        throw this._handleError(error_3);
                    case 8: return [2 /*return*/];
                }
            });
        });
    };
    GroupMembershipService.prototype.getUserPersonalSite = function () {
        return __awaiter(this, void 0, void 0, function () {
            var response, error, error_5;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        if (this._userSite) {
                            return [2 /*return*/, this._userSite];
                        }
                        return [4 /*yield*/, this.graphClient
                                .api('/me/drive')
                                .select('id,name,webUrl')
                                .get()];
                    case 1:
                        response = _a.sent();
                        if (!response) {
                            error = {
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
                        return [2 /*return*/, this._userSite];
                    case 2:
                        error_5 = _a.sent();
                        throw this._handleError(error_5);
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    GroupMembershipService.prototype.syncFilesToUserSite = function () {
        return __awaiter(this, void 0, void 0, function () {
            var userFiles, syncFolderName, syncFolderId, existingFolder, _a, newFolder, teamGroups, totalSynced, _i, _b, teamName, files, teamFolderId, existingTeamFolder, _c, newTeamFolder, _d, files_1, file, copyError_1, teamError_2, error_6;
            return __generator(this, function (_e) {
                switch (_e.label) {
                    case 0:
                        _e.trys.push([0, 25, , 26]);
                        this._syncStatus.status = 'syncing';
                        this._syncStatus.error = undefined;
                        return [4 /*yield*/, this.getUserPersonalSite()];
                    case 1:
                        _e.sent();
                        return [4 /*yield*/, this.getUserFilesInTeams()];
                    case 2:
                        userFiles = _e.sent();
                        if (userFiles.length === 0) {
                            this._syncStatus.status = 'idle';
                            return [2 /*return*/];
                        }
                        return [4 /*yield*/, this.graphClient
                                .api('/me/drive')
                                .get()];
                    case 3:
                        _e.sent();
                        syncFolderName = 'Teams File Sync';
                        syncFolderId = void 0;
                        _e.label = 4;
                    case 4:
                        _e.trys.push([4, 6, , 8]);
                        return [4 /*yield*/, this.graphClient
                                .api("/me/drive/root:/".concat(syncFolderName))
                                .get()];
                    case 5:
                        existingFolder = _e.sent();
                        syncFolderId = existingFolder.id;
                        return [3 /*break*/, 8];
                    case 6:
                        _a = _e.sent();
                        return [4 /*yield*/, this.graphClient
                                .api('/me/drive/root/children')
                                .post({
                                name: syncFolderName,
                                folder: {},
                                '@microsoft.graph.conflictBehavior': 'replace'
                            })];
                    case 7:
                        newFolder = _e.sent();
                        syncFolderId = newFolder.id;
                        return [3 /*break*/, 8];
                    case 8:
                        teamGroups = this._groupFilesByTeam(userFiles);
                        totalSynced = 0;
                        _i = 0, _b = Object.keys(teamGroups);
                        _e.label = 9;
                    case 9:
                        if (!(_i < _b.length)) return [3 /*break*/, 24];
                        teamName = _b[_i];
                        files = teamGroups[teamName];
                        _e.label = 10;
                    case 10:
                        _e.trys.push([10, 22, , 23]);
                        teamFolderId = void 0;
                        _e.label = 11;
                    case 11:
                        _e.trys.push([11, 13, , 15]);
                        return [4 /*yield*/, this.graphClient
                                .api("/me/drive/items/".concat(syncFolderId, ":/").concat(teamName))
                                .get()];
                    case 12:
                        existingTeamFolder = _e.sent();
                        teamFolderId = existingTeamFolder.id;
                        return [3 /*break*/, 15];
                    case 13:
                        _c = _e.sent();
                        return [4 /*yield*/, this.graphClient
                                .api("/me/drive/items/".concat(syncFolderId, "/children"))
                                .post({
                                name: teamName,
                                folder: {},
                                '@microsoft.graph.conflictBehavior': 'replace'
                            })];
                    case 14:
                        newTeamFolder = _e.sent();
                        teamFolderId = newTeamFolder.id;
                        return [3 /*break*/, 15];
                    case 15:
                        _d = 0, files_1 = files;
                        _e.label = 16;
                    case 16:
                        if (!(_d < files_1.length)) return [3 /*break*/, 21];
                        file = files_1[_d];
                        _e.label = 17;
                    case 17:
                        _e.trys.push([17, 19, , 20]);
                        return [4 /*yield*/, this._copyFileToUserSite(file, teamFolderId)];
                    case 18:
                        _e.sent();
                        totalSynced++;
                        return [3 /*break*/, 20];
                    case 19:
                        copyError_1 = _e.sent();
                        console.warn("Failed to copy file ".concat(file.name, ":"), copyError_1);
                        return [3 /*break*/, 20];
                    case 20:
                        _d++;
                        return [3 /*break*/, 16];
                    case 21: return [3 /*break*/, 23];
                    case 22:
                        teamError_2 = _e.sent();
                        console.warn("Failed to sync files for team ".concat(teamName, ":"), teamError_2);
                        return [3 /*break*/, 23];
                    case 23:
                        _i++;
                        return [3 /*break*/, 9];
                    case 24:
                        this._syncStatus = {
                            isEnabled: true,
                            lastSyncDate: new Date(),
                            fileCount: totalSynced,
                            teamsCount: Object.keys(teamGroups).length,
                            status: 'idle'
                        };
                        this._saveSyncStatus();
                        return [3 /*break*/, 26];
                    case 25:
                        error_6 = _e.sent();
                        this._syncStatus.status = 'error';
                        this._syncStatus.error = 'Sync failed: ' + error_6.message;
                        throw this._handleError(error_6, 'SYNC_ERROR');
                    case 26: return [2 /*return*/];
                }
            });
        });
    };
    GroupMembershipService.prototype.getSyncStatus = function () {
        return __assign({}, this._syncStatus);
    };
    GroupMembershipService.prototype.toggleSync = function (enabled) {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        this._syncStatus.isEnabled = enabled;
                        this._saveSyncPreference(enabled);
                        if (!enabled) return [3 /*break*/, 2];
                        return [4 /*yield*/, this.syncFilesToUserSite()];
                    case 1:
                        _a.sent();
                        _a.label = 2;
                    case 2: return [2 /*return*/];
                }
            });
        });
    };
    GroupMembershipService.prototype.startBackgroundSync = function () {
        return __awaiter(this, void 0, void 0, function () {
            var error_7;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (!this._syncStatus.isEnabled)
                            return [2 /*return*/];
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, this.syncFilesToUserSite()];
                    case 2:
                        _a.sent();
                        return [3 /*break*/, 4];
                    case 3:
                        error_7 = _a.sent();
                        console.error('Background sync failed:', error_7);
                        return [3 /*break*/, 4];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    GroupMembershipService.prototype._getCurrentUser = function () {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.graphClient
                            .api('/me')
                            .select('id,displayName,mail,userPrincipalName')
                            .get()];
                    case 1: return [2 /*return*/, _a.sent()];
                }
            });
        });
    };
    GroupMembershipService.prototype._getAllFilesFromDrive = function (driveId, itemId, path) {
        if (path === void 0) { path = ''; }
        return __awaiter(this, void 0, void 0, function () {
            var response, allFiles, _i, _a, item, subFiles, error_8;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        _b.trys.push([0, 7, , 8]);
                        return [4 /*yield*/, this.graphClient
                                .api("/drives/".concat(driveId, "/items/").concat(itemId, "/children"))
                                .select('id,name,size,webUrl,lastModifiedDateTime,lastModifiedBy,createdBy,parentReference,file,folder')
                                .get()];
                    case 1:
                        response = _b.sent();
                        if (!response || !response.value)
                            return [2 /*return*/, []];
                        allFiles = [];
                        _i = 0, _a = response.value;
                        _b.label = 2;
                    case 2:
                        if (!(_i < _a.length)) return [3 /*break*/, 6];
                        item = _a[_i];
                        if (!item.file) return [3 /*break*/, 3];
                        item.parentReference = __assign(__assign({}, item.parentReference), { path: path });
                        allFiles.push(item);
                        return [3 /*break*/, 5];
                    case 3:
                        if (!item.folder) return [3 /*break*/, 5];
                        return [4 /*yield*/, this._getAllFilesFromDrive(driveId, item.id, "".concat(path, "/").concat(item.name))];
                    case 4:
                        subFiles = _b.sent();
                        allFiles.push.apply(allFiles, subFiles);
                        _b.label = 5;
                    case 5:
                        _i++;
                        return [3 /*break*/, 2];
                    case 6: return [2 /*return*/, allFiles];
                    case 7:
                        error_8 = _b.sent();
                        console.warn("Failed to get files from drive ".concat(driveId, ", item ").concat(itemId, ":"), error_8);
                        return [2 /*return*/, []];
                    case 8: return [2 /*return*/];
                }
            });
        });
    };
    GroupMembershipService.prototype._copyFileToUserSite = function (file, destinationFolderId) {
        return __awaiter(this, void 0, void 0, function () {
            var sourceUrl, copyResponse, _a, _b, error_9;
            var _c, _d;
            return __generator(this, function (_e) {
                switch (_e.label) {
                    case 0:
                        _e.trys.push([0, 3, , 4]);
                        sourceUrl = "/drives/".concat(file.driveId, "/items/").concat(file.id);
                        _b = (_a = this.graphClient
                            .api("".concat(sourceUrl, "/copy")))
                            .post;
                        _c = {};
                        _d = {};
                        return [4 /*yield*/, this.graphClient.api('/me/drive').get()];
                    case 1: return [4 /*yield*/, _b.apply(_a, [(_c.parentReference = (_d.driveId = (_e.sent()).id,
                                _d.id = destinationFolderId,
                                _d),
                                _c.name = file.name,
                                _c)])];
                    case 2:
                        copyResponse = _e.sent();
                        if (copyResponse && copyResponse.id) {
                            console.log("Successfully copied file: ".concat(file.name));
                        }
                        return [3 /*break*/, 4];
                    case 3:
                        error_9 = _e.sent();
                        if (error_9.code !== 'nameAlreadyExists') {
                            throw error_9;
                        }
                        return [3 /*break*/, 4];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    GroupMembershipService.prototype._groupFilesByTeam = function (files) {
        return files.reduce(function (groups, file) {
            if (!groups[file.teamName]) {
                groups[file.teamName] = [];
            }
            groups[file.teamName].push(file);
            return groups;
        }, {});
    };
    GroupMembershipService.prototype._getFileType = function (fileName) {
        var _a;
        var extension = ((_a = fileName.split('.').pop()) === null || _a === void 0 ? void 0 : _a.toLowerCase()) || '';
        var types = {
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
    };
    GroupMembershipService.prototype._getSyncPreference = function () {
        try {
            var stored = localStorage.getItem('teamsFileSyncEnabled');
            return stored === 'true';
        }
        catch (_a) {
            return false;
        }
    };
    GroupMembershipService.prototype._saveSyncPreference = function (enabled) {
        try {
            localStorage.setItem('teamsFileSyncEnabled', enabled.toString());
        }
        catch (error) {
            console.warn('Failed to save sync preference:', error);
        }
    };
    GroupMembershipService.prototype._saveSyncStatus = function () {
        var _a;
        try {
            var statusToSave = {
                lastSyncDate: (_a = this._syncStatus.lastSyncDate) === null || _a === void 0 ? void 0 : _a.toISOString(),
                fileCount: this._syncStatus.fileCount,
                teamsCount: this._syncStatus.teamsCount
            };
            localStorage.setItem('teamsFileSyncStatus', JSON.stringify(statusToSave));
        }
        catch (error) {
            console.warn('Failed to save sync status:', error);
        }
    };
    GroupMembershipService.prototype._loadSyncStatus = function () {
        try {
            var stored = localStorage.getItem('teamsFileSyncStatus');
            if (stored) {
                var parsed = JSON.parse(stored);
                this._syncStatus.lastSyncDate = parsed.lastSyncDate ? new Date(parsed.lastSyncDate) : undefined;
                this._syncStatus.fileCount = parsed.fileCount || 0;
                this._syncStatus.teamsCount = parsed.teamsCount || 0;
            }
        }
        catch (error) {
            console.warn('Failed to load sync status:', error);
        }
    };
    GroupMembershipService.prototype._handleError = function (error, defaultType) {
        if (defaultType === void 0) { defaultType = 'UNKNOWN'; }
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
        if (error.type) {
            return error;
        }
        return {
            type: defaultType,
            message: "An error occurred: ".concat(error.message || 'Unknown error'),
            originalError: error
        };
    };
    return GroupMembershipService;
}());
export { GroupMembershipService };
//# sourceMappingURL=GroupMembershipService.js.map