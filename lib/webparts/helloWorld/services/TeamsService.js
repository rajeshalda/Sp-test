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
var TeamsService = /** @class */ (function () {
    function TeamsService(graphClient) {
        this.graphClient = graphClient;
    }
    TeamsService.prototype.getUserTeams = function () {
        return __awaiter(this, void 0, void 0, function () {
            var response, error_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, this.graphClient
                                .api('/me/joinedTeams')
                                .get()];
                    case 1:
                        response = _a.sent();
                        if (!response || !response.value) {
                            throw new Error('No teams data received');
                        }
                        return [2 /*return*/, response.value.map(function (team) { return ({
                                id: team.id,
                                displayName: team.displayName,
                                description: team.description
                            }); })];
                    case 2:
                        error_1 = _a.sent();
                        throw this._handleError(error_1);
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    TeamsService.prototype.getTeamChannels = function (teamId) {
        return __awaiter(this, void 0, void 0, function () {
            var response, error_2;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, this.graphClient
                                .api("/teams/".concat(teamId, "/channels"))
                                .get()];
                    case 1:
                        response = _a.sent();
                        if (!response || !response.value) {
                            throw new Error('No channels data received');
                        }
                        return [2 /*return*/, response.value.map(function (channel) { return ({
                                id: channel.id,
                                displayName: channel.displayName,
                                description: channel.description
                            }); })];
                    case 2:
                        error_2 = _a.sent();
                        throw this._handleError(error_2);
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    TeamsService.prototype.getAllChannelsForUser = function () {
        return __awaiter(this, void 0, void 0, function () {
            var teams, error, allChannels, _loop_1, this_1, _i, teams_1, team, error_3;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 6, , 7]);
                        return [4 /*yield*/, this.getUserTeams()];
                    case 1:
                        teams = _a.sent();
                        if (teams.length === 0) {
                            error = {
                                type: 'NO_TEAMS',
                                message: 'User is not a member of any teams'
                            };
                            throw error;
                        }
                        allChannels = [];
                        _loop_1 = function (team) {
                            var channels, channelsWithTeamInfo, error_4;
                            return __generator(this, function (_b) {
                                switch (_b.label) {
                                    case 0:
                                        _b.trys.push([0, 2, , 3]);
                                        return [4 /*yield*/, this_1.getTeamChannels(team.id)];
                                    case 1:
                                        channels = _b.sent();
                                        channelsWithTeamInfo = channels.map(function (channel) { return (__assign(__assign({}, channel), { teamId: team.id, teamName: team.displayName })); });
                                        allChannels.push.apply(allChannels, channelsWithTeamInfo);
                                        return [3 /*break*/, 3];
                                    case 2:
                                        error_4 = _b.sent();
                                        console.warn("Failed to get channels for team ".concat(team.displayName, ":"), error_4);
                                        return [3 /*break*/, 3];
                                    case 3: return [2 /*return*/];
                                }
                            });
                        };
                        this_1 = this;
                        _i = 0, teams_1 = teams;
                        _a.label = 2;
                    case 2:
                        if (!(_i < teams_1.length)) return [3 /*break*/, 5];
                        team = teams_1[_i];
                        return [5 /*yield**/, _loop_1(team)];
                    case 3:
                        _a.sent();
                        _a.label = 4;
                    case 4:
                        _i++;
                        return [3 /*break*/, 2];
                    case 5: return [2 /*return*/, allChannels];
                    case 6:
                        error_3 = _a.sent();
                        if (error_3.type) {
                            throw error_3;
                        }
                        throw this._handleError(error_3);
                    case 7: return [2 /*return*/];
                }
            });
        });
    };
    TeamsService.prototype.getChannelFilesFolder = function (teamId, channelId) {
        return __awaiter(this, void 0, void 0, function () {
            var response, error_5;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, this.graphClient
                                .api("/teams/".concat(teamId, "/channels/").concat(channelId, "/filesFolder"))
                                .get()];
                    case 1:
                        response = _a.sent();
                        console.log('Raw filesFolder response:', JSON.stringify(response, null, 2));
                        return [2 /*return*/, {
                                id: response.id,
                                name: response.name,
                                webUrl: response.webUrl,
                                size: response.size,
                                folder: response.folder,
                                createdDateTime: response.createdDateTime,
                                lastModifiedDateTime: response.lastModifiedDateTime
                            }];
                    case 2:
                        error_5 = _a.sent();
                        throw this._handleError(error_5);
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    TeamsService.prototype.getDriveItemChildren = function (driveId, itemId) {
        if (itemId === void 0) { itemId = 'root'; }
        return __awaiter(this, void 0, void 0, function () {
            var response, error_6;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, this.graphClient
                                .api("/drives/".concat(driveId, "/items/").concat(itemId, "/children"))
                                .get()];
                    case 1:
                        response = _a.sent();
                        if (!response || !response.value) {
                            return [2 /*return*/, []];
                        }
                        return [2 /*return*/, response.value.map(function (item) { return ({
                                id: item.id,
                                name: item.name,
                                webUrl: item.webUrl,
                                size: item.size,
                                folder: item.folder,
                                file: item.file,
                                createdDateTime: item.createdDateTime,
                                lastModifiedDateTime: item.lastModifiedDateTime
                            }); })];
                    case 2:
                        error_6 = _a.sent();
                        console.warn('Failed to get drive item children:', error_6);
                        return [2 /*return*/, []];
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    TeamsService.prototype.buildTeamsTreeView = function () {
        return __awaiter(this, void 0, void 0, function () {
            var teams, treeNodes, _i, teams_2, team, teamNode, channels, _a, channels_1, channel, channelNode, error_7, error_8;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        _b.trys.push([0, 9, , 10]);
                        return [4 /*yield*/, this.getUserTeams()];
                    case 1:
                        teams = _b.sent();
                        treeNodes = [];
                        _i = 0, teams_2 = teams;
                        _b.label = 2;
                    case 2:
                        if (!(_i < teams_2.length)) return [3 /*break*/, 8];
                        team = teams_2[_i];
                        teamNode = {
                            id: team.id,
                            name: team.displayName,
                            type: 'team',
                            description: team.description,
                            isExpanded: false,
                            children: []
                        };
                        _b.label = 3;
                    case 3:
                        _b.trys.push([3, 5, , 6]);
                        return [4 /*yield*/, this.getTeamChannels(team.id)];
                    case 4:
                        channels = _b.sent();
                        for (_a = 0, channels_1 = channels; _a < channels_1.length; _a++) {
                            channel = channels_1[_a];
                            channelNode = {
                                id: "".concat(team.id, "_").concat(channel.id),
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
                        return [3 /*break*/, 6];
                    case 5:
                        error_7 = _b.sent();
                        console.warn("Failed to get channels for team ".concat(team.displayName, ":"), error_7);
                        return [3 /*break*/, 6];
                    case 6:
                        treeNodes.push(teamNode);
                        _b.label = 7;
                    case 7:
                        _i++;
                        return [3 /*break*/, 2];
                    case 8: return [2 /*return*/, treeNodes];
                    case 9:
                        error_8 = _b.sent();
                        throw this._handleError(error_8);
                    case 10: return [2 /*return*/];
                }
            });
        });
    };
    TeamsService.prototype.loadChannelFiles = function (teamId, channelId) {
        return __awaiter(this, void 0, void 0, function () {
            var filesFolder, driveResponse, rootItems, channelFolderItems, _i, rootItems_1, item, folderContents, folderData, folderError_1, error_9;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 13, , 14]);
                        return [4 /*yield*/, this.getChannelFilesFolder(teamId, channelId)];
                    case 1:
                        filesFolder = _a.sent();
                        console.log('FilesFolder response:', filesFolder);
                        // Extract the drive and item information from the filesFolder
                        // The filesFolder should contain drive information
                        if (!filesFolder.webUrl) {
                            console.warn('No webUrl in filesFolder response');
                            return [2 /*return*/, []];
                        }
                        return [4 /*yield*/, this.graphClient
                                .api("/groups/".concat(teamId, "/drive"))
                                .get()];
                    case 2:
                        driveResponse = _a.sent();
                        if (!driveResponse) {
                            console.warn('No drive response for team');
                            return [2 /*return*/, []];
                        }
                        console.log('Drive response:', driveResponse);
                        return [4 /*yield*/, this.getDriveItemChildren(driveResponse.id, 'root')];
                    case 3:
                        rootItems = _a.sent();
                        console.log('Root items:', rootItems);
                        channelFolderItems = [];
                        _i = 0, rootItems_1 = rootItems;
                        _a.label = 4;
                    case 4:
                        if (!(_i < rootItems_1.length)) return [3 /*break*/, 7];
                        item = rootItems_1[_i];
                        if (!item.folder) return [3 /*break*/, 6];
                        return [4 /*yield*/, this.getDriveItemChildren(driveResponse.id, item.id)];
                    case 5:
                        folderContents = _a.sent();
                        console.log("Contents of ".concat(item.name, ":"), folderContents);
                        channelFolderItems = channelFolderItems.concat(folderContents);
                        _a.label = 6;
                    case 6:
                        _i++;
                        return [3 /*break*/, 4];
                    case 7:
                        if (!(channelFolderItems.length === 0)) return [3 /*break*/, 12];
                        _a.label = 8;
                    case 8:
                        _a.trys.push([8, 11, , 12]);
                        folderData = filesFolder;
                        if (!(folderData.parentReference && folderData.parentReference.driveId)) return [3 /*break*/, 10];
                        return [4 /*yield*/, this.getDriveItemChildren(folderData.parentReference.driveId, folderData.id)];
                    case 9:
                        channelFolderItems = _a.sent();
                        _a.label = 10;
                    case 10: return [3 /*break*/, 12];
                    case 11:
                        folderError_1 = _a.sent();
                        console.warn('Fallback method failed:', folderError_1);
                        return [3 /*break*/, 12];
                    case 12: return [2 /*return*/, this._convertDriveItemsToTreeNodes(channelFolderItems, driveResponse.id, "".concat(teamId, "_").concat(channelId))];
                    case 13:
                        error_9 = _a.sent();
                        console.error("Failed to load files for channel ".concat(channelId, ":"), error_9);
                        return [2 /*return*/, []];
                    case 14: return [2 /*return*/];
                }
            });
        });
    };
    TeamsService.prototype.loadChannelFilesDirect = function (teamId, channelId) {
        return __awaiter(this, void 0, void 0, function () {
            var filesFolder, fullResponse, children, error_10;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 5, , 6]);
                        console.log("\uD83D\uDD0D Starting loadChannelFilesDirect for teamId: ".concat(teamId, ", channelId: ").concat(channelId));
                        return [4 /*yield*/, this.getChannelFilesFolder(teamId, channelId)];
                    case 1:
                        filesFolder = _a.sent();
                        console.log('📁 FilesFolder data:', JSON.stringify(filesFolder, null, 2));
                        if (!filesFolder.id) return [3 /*break*/, 4];
                        return [4 /*yield*/, this.graphClient
                                .api("/teams/".concat(teamId, "/channels/").concat(channelId, "/filesFolder"))
                                .select('id,name,parentReference,webUrl')
                                .get()];
                    case 2:
                        fullResponse = _a.sent();
                        console.log('🔧 Full filesFolder response:', JSON.stringify(fullResponse, null, 2));
                        if (!(fullResponse.parentReference && fullResponse.parentReference.driveId)) return [3 /*break*/, 4];
                        console.log("\uD83D\uDCC2 Found drive ID: ".concat(fullResponse.parentReference.driveId, ", item ID: ").concat(fullResponse.id));
                        return [4 /*yield*/, this.getDriveItemChildren(fullResponse.parentReference.driveId, fullResponse.id)];
                    case 3:
                        children = _a.sent();
                        console.log("\uD83D\uDCC4 Found ".concat(children.length, " items in channel folder"));
                        if (children.length > 0) {
                            console.log('✅ Successfully loaded files:', children.map(function (c) { return c.name; }));
                            return [2 /*return*/, this._convertDriveItemsToTreeNodes(children, fullResponse.parentReference.driveId, "".concat(teamId, "_").concat(channelId))];
                        }
                        _a.label = 4;
                    case 4:
                        console.log('⚠️ No files found or no drive reference available');
                        return [2 /*return*/, []];
                    case 5:
                        error_10 = _a.sent();
                        console.error("\u274C Direct approach failed for channel ".concat(channelId, ":"), error_10);
                        console.error('Error details:', JSON.stringify(error_10, null, 2));
                        return [2 /*return*/, []];
                    case 6: return [2 /*return*/];
                }
            });
        });
    };
    TeamsService.prototype._convertDriveItemsToTreeNodes = function (driveItems, driveId, parentId) {
        return driveItems.map(function (item) {
            var _a;
            return ({
                id: "".concat(parentId, "_").concat(item.id),
                name: item.name,
                type: item.folder ? 'folder' : 'file',
                parentId: parentId,
                webUrl: item.webUrl,
                size: item.size,
                driveId: driveId,
                itemId: item.id,
                childCount: ((_a = item.folder) === null || _a === void 0 ? void 0 : _a.childCount) || 0,
                isExpanded: false,
                children: [],
                isLoading: false
            });
        });
    };
    TeamsService.prototype._handleError = function (error) {
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
            message: "An unexpected error occurred: ".concat(error.message || 'Unknown error'),
            originalError: error
        };
    };
    return TeamsService;
}());
export { TeamsService };
//# sourceMappingURL=TeamsService.js.map