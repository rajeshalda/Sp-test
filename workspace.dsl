workspace "DASL - Infra Central Library (ICL)" "C4 Model for SPFx Infra Central Library Web Part" {

    model {
        # Define people/users
        user = person "SharePoint User" "A user who needs to sync files from Teams to their personal SharePoint site" "User"
        administrator = person "SharePoint Administrator" "Manages SharePoint and approves Microsoft Graph API permissions" "Administrator"

        # Define external systems
        microsoft365 = softwareSystem "Microsoft 365" "Microsoft cloud platform providing identity, collaboration, and storage services" "External System" {
            azureAd = container "Azure Active Directory" "Provides authentication and authorization services" "Identity Provider" "External"
            microsoftGraph = container "Microsoft Graph API" "Unified API endpoint for Microsoft 365 services" "REST API" "External"
            teamsBackend = container "Microsoft Teams" "Team collaboration platform with file storage" "SaaS Platform" "External"
            sharepointOnline = container "SharePoint Online" "Document management and storage platform" "SaaS Platform" "External"
        }

        # Define the main system
        fileSyncSystem = softwareSystem "Infra Central Library (ICL)" "Enables users to automatically sync files from Microsoft Teams to their personal SharePoint site" {

            # Container: SPFx Web Part
            webPart = container "ICL Web Part" "SharePoint Framework client-side web part that provides the user interface and orchestrates file sync operations" "TypeScript, SPFx 1.21.1" {

                # Components within the Web Part
                uiComponent = component "UI Renderer" "Renders the user interface with sync controls, status display, and team listings" "React-style rendering"
                eventHandler = component "Event Handler" "Handles user interactions (clicks, form submissions) and triggers appropriate actions" "Event Listeners"
                stateManager = component "State Manager" "Manages application state including sync status, teams list, and user files" "State Management"
                syncController = component "Sync Controller" "Coordinates sync initialization, toggling, and background sync operations" "Controller"
                errorHandler = component "Error Handler" "Handles and displays user-friendly error messages for different error types" "Error Handling"

                groupService = component "GroupMembershipService" "Service layer that interacts with Microsoft Graph to retrieve teams, files, and manage sync operations" "TypeScript Service Class" "InProgress"

                spfxContext = component "SPFx Context" "Provides access to SharePoint Framework APIs, user context, and Graph client factory" "SPFx API" "InProgress"
            }

            # Container: Background Sync Process
            backgroundSync = container "Background Sync Process" "Automated process that periodically syncs files from Teams to SharePoint" "JavaScript Timer" {
                syncScheduler = component "Sync Scheduler" "Schedules and triggers sync operations every 4 hours when sync is enabled" "setInterval Timer"
                syncExecutor = component "Sync Executor" "Executes the actual file sync logic by calling the service layer" "Async Task"
            }
        }

        # Define the On-Premises Download Service (Separate System)
        downloadService = softwareSystem "Teams File Download Service" "On-premises Python service that downloads files from SharePoint to local storage for RAG processing with delta sync capability" {

            # Container: Download Service Application
            downloadApp = container "Download Service Application" "Python application that orchestrates file download from SharePoint to on-premise storage" "Python 3.8+, MSAL, Microsoft Graph API" {
                authModule = component "Azure AD Authenticator" "Handles Azure AD authentication using MSAL with client credentials flow" "Python/MSAL" "Completed"
                graphClient = component "Graph API Client" "Manages Microsoft Graph API interactions with retry logic and error handling" "Python/Requests" "Completed"
                fileDownloader = component "File Downloader" "Coordinates file discovery, delta sync, and download operations" "Python" "Completed"
                deltaSync = component "Delta Sync Tracker" "Tracks downloaded files using JSON state file to avoid re-downloading unchanged files" "JSON State Manager" "Completed"
                mainOrchestrator = component "Main Orchestrator" "Entry point that orchestrates authentication, discovery, and download workflow" "Python" "Completed"
                logger = component "Logging System" "Dual logging to console and file with comprehensive error tracking" "Python Logging" "Completed"
            }

            # Container: Local File Storage
            localStorage = container "Local File Storage" "File system storage for downloaded files organized by Team folders" "File System" {
                description "Stores downloaded files with preserved folder structure and sync state JSON file"
            }
        }

        # Define relationships between people and systems
        user -> fileSyncSystem "Uses to sync Teams files to personal SharePoint"
        user -> microsoft365 "Authenticates with"
        administrator -> azureAd "Approves API permissions"
        administrator -> fileSyncSystem "Deploys and configures"
        administrator -> downloadService "Deploys and schedules on-premises"

        # System-to-system relationships
        fileSyncSystem -> microsoft365 "Retrieves teams, files, and performs sync operations using"
        fileSyncSystem -> azureAd "Authenticates users via"
        downloadService -> microsoft365 "Downloads files from SharePoint via"
        downloadService -> azureAd "Authenticates with application permissions via"
        downloadService -> sharepointOnline "Reads files from Teams File Sync folder"

        # Container-to-container relationships (ICL System)
        webPart -> backgroundSync "Controls and monitors"
        webPart -> microsoftGraph "Queries teams and files via"
        webPart -> sharepointOnline "Creates folders and copies files to"
        webPart -> azureAd "Authenticates via"
        backgroundSync -> microsoftGraph "Periodically syncs files via"
        backgroundSync -> sharepointOnline "Copies files to"

        # Container-to-container relationships (Download Service System)
        downloadApp -> azureAd "Authenticates with application permissions via"
        downloadApp -> microsoftGraph "Downloads files via"
        downloadApp -> sharepointOnline "Reads files from Teams File Sync folder"
        downloadApp -> localStorage "Saves downloaded files to"

        # Component relationships within Web Part
        uiComponent -> stateManager "Reads state from"
        eventHandler -> syncController "Triggers sync operations"
        eventHandler -> stateManager "Updates state"
        syncController -> groupService "Delegates Graph API calls to"
        syncController -> errorHandler "Reports errors to"
        syncController -> stateManager "Updates sync status in"
        groupService -> spfxContext "Gets Graph client from"
        groupService -> microsoftGraph "Makes API calls to"
        errorHandler -> uiComponent "Displays errors via"
        stateManager -> uiComponent "Notifies of state changes"

        # Component to external system relationships
        spfxContext -> azureAd "Obtains access tokens from"
        groupService -> teamsBackend "Retrieves connected teams from"
        groupService -> sharepointOnline "Gets user's personal site and creates sync folders"

        # Background sync relationships
        syncScheduler -> syncExecutor "Triggers sync every 4 hours"
        syncController -> syncScheduler "Starts/stops"

        # Additional component relationships
        uiComponent -> eventHandler "Captures user interactions from"
        spfxContext -> groupService "Provides Graph client to"

        # Download Service component relationships (within downloadApp container)
        mainOrchestrator -> authModule "Initiates authentication"
        mainOrchestrator -> graphClient "Provides authenticated client to"
        mainOrchestrator -> fileDownloader "Orchestrates download process via"
        mainOrchestrator -> logger "Logs operations via"
        authModule -> mainOrchestrator "Returns access token to"
        fileDownloader -> graphClient "Uses to discover and download files"
        fileDownloader -> deltaSync "Checks sync state and updates"
        fileDownloader -> logger "Logs progress and errors to"
        graphClient -> logger "Logs API responses to"
        deltaSync -> fileDownloader "Provides file state information to"

        # Download Service to external systems
        authModule -> azureAd "Requests access token from"
        graphClient -> microsoftGraph "Makes API calls to"
        fileDownloader -> localStorage "Writes downloaded files to"
        deltaSync -> localStorage "Reads/writes sync_state.json to"

        # Deployment
        deploymentEnvironment = deploymentEnvironment "Production" {
            deploymentNode "User's Browser" "Web Browser" "Chrome, Edge, Safari" {
                deploymentNode "SharePoint Page" "SharePoint Modern Page" "SharePoint Online" {
                    webPartInstance = containerInstance webPart
                    backgroundSyncInstance = containerInstance backgroundSync
                }
            }

            deploymentNode "On-Premise Server" "Windows/Linux Server" "Local Infrastructure" {
                deploymentNode "Python Runtime" "Python 3.8+" "Runtime Environment" {
                    downloadAppInstance = containerInstance downloadApp
                }
                localStorageInstance = containerInstance localStorage
            }

            deploymentNode "Microsoft Cloud" "Microsoft Azure" "Cloud Infrastructure" {
                deploymentNode "Microsoft 365 Tenant" "" "Multi-tenant SaaS" {
                    azureAdInstance = containerInstance azureAd
                    graphInstance = containerInstance microsoftGraph
                    teamsInstance = containerInstance teamsBackend
                    sharepointInstance = containerInstance sharepointOnline
                }
            }
        }
    }

    views {
        # COMPREHENSIVE OVERVIEW - All-in-one diagram showing both systems
        systemLandscape "ComprehensiveOverview" {
            include *
            autoLayout
            description "Comprehensive overview showing ICL system, on-premises download service, users, and Microsoft 365 integration"
        }

        # System Context diagram
        systemContext fileSyncSystem "SystemContext" {
            include *
            autoLayout
            description "System Context diagram for Teams File Sync System showing how users and administrators interact with the system and Microsoft 365 services"
        }

        # Container diagram
        container fileSyncSystem "Containers" {
            include *
            autoLayout
            description "Container diagram showing the main containers of the Teams File Sync System"
        }

        # Component diagram for Web Part
        component webPart "WebPartComponents" {
            include *
            autoLayout
            description "Component diagram showing the internal structure of the HelloWorld Web Part"
        }

        # Component diagram for Background Sync
        component backgroundSync "BackgroundSyncComponents" {
            include *
            autoLayout
            description "Component diagram showing the background sync process components"
        }

        # System Context for Download Service
        systemContext downloadService "DownloadServiceContext" {
            include *
            autoLayout
            description "System Context diagram for On-Premises Teams File Download Service"
        }

        # Container diagram for Download Service
        container downloadService "DownloadServiceContainers" {
            include *
            autoLayout
            description "Container diagram showing the Download Service application and local storage"
        }

        # Component diagram for Download Service
        component downloadApp "DownloadServiceComponents" {
            include *
            autoLayout
            description "Component diagram showing the Teams File Download Service components and workflow"
        }

        # Deployment diagram
        deployment fileSyncSystem "Production" "DeploymentProduction" {
            include *
            autoLayout
            description "Deployment diagram showing how the system is deployed in production"
        }

        # Dynamic diagram - Initialize Sync Flow
        dynamic webPart "InitializeSyncFlow" "Illustrates the flow when a user initializes file sync" {
            uiComponent -> eventHandler "User clicks 'Initialize File Sync' button"
            eventHandler -> syncController "Calls _initializeSync()"
            syncController -> stateManager "Sets loading state"
            stateManager -> uiComponent "Updates UI to show loading spinner"
            syncController -> groupService "Calls getUserPersonalSite()"
            groupService -> spfxContext "Gets Graph client"
            spfxContext -> groupService "Returns Graph client instance"
            groupService -> spfxContext "Requests access token"
            spfxContext -> groupService "Returns token"
            syncController -> groupService "Calls getConnectedTeams()"
            syncController -> groupService "Calls getUserFilesInTeams()"
            syncController -> groupService "Calls getSyncStatus()"
            groupService -> syncController "Returns current sync status"
            syncController -> stateManager "Updates state with teams and files"
            stateManager -> uiComponent "Triggers re-render with sync interface"
            autoLayout
            description "Sequence showing the initialization of file sync when user clicks the button"
        }

        # Dynamic diagram - Enable Sync Flow
        dynamic webPart "EnableSyncFlow" "Illustrates what happens when sync is enabled" {
            uiComponent -> eventHandler "User clicks 'Enable Sync' button"
            eventHandler -> syncController "Calls _toggleSync(true)"
            syncController -> stateManager "Sets loading state"
            syncController -> groupService "Calls toggleSync(true)"
            groupService -> syncController "Returns success"
            syncController -> groupService "Calls getSyncStatus()"
            groupService -> syncController "Returns updated status"
            syncController -> syncScheduler "Calls _startBackgroundSync()"
            syncScheduler -> syncExecutor "Initializes sync executor"
            syncController -> stateManager "Updates sync status"
            stateManager -> uiComponent "Re-renders with enabled state"
            autoLayout
            description "Sequence showing what happens when user enables file sync"
        }

        # Dynamic diagram - File Download Flow (NEW)
        dynamic downloadApp "FileDownloadFlow" "Illustrates the complete file download workflow from SharePoint to on-premise storage" {
            mainOrchestrator -> logger "Logs: Starting file download service"
            mainOrchestrator -> authModule "Initiates Azure AD authentication"
            authModule -> mainOrchestrator "Returns access token"
            mainOrchestrator -> graphClient "Creates authenticated Graph client"
            mainOrchestrator -> fileDownloader "Starts file discovery and download"
            fileDownloader -> graphClient "Calls getUserDrive()"
            fileDownloader -> graphClient "Calls getFolderByPath('Teams File Sync')"
            fileDownloader -> graphClient "Recursively discovers all files"
            fileDownloader -> deltaSync "Loads sync state from JSON"
            deltaSync -> fileDownloader "Returns list of previously downloaded files"
            fileDownloader -> deltaSync "Filters new and modified files"
            fileDownloader -> logger "Logs: Found X files, Y need download"
            fileDownloader -> graphClient "Downloads each file with progress tracking"
            graphClient -> fileDownloader "Streams file content"
            fileDownloader -> deltaSync "Updates sync state with downloaded files"
            fileDownloader -> logger "Logs: Sync state saved to disk"
            fileDownloader -> mainOrchestrator "Returns download statistics"
            mainOrchestrator -> logger "Logs: Download complete summary"
            autoLayout
            description "Sequence showing the complete file download workflow including delta sync"
        }

        # Styling
        styles {
            element "Person" {
                shape Person
                background #08427b
                color #ffffff
            }
            element "User" {
                background #1168bd
            }
            element "Administrator" {
                background #999999
            }
            element "Software System" {
                background #1168bd
                color #ffffff
            }
            element "External System" {
                background #999999
                color #ffffff
            }
            element "Container" {
                background #438dd5
                color #ffffff
            }
            element "External" {
                background #cccccc
                color #000000
            }
            element "Component" {
                background #85bbf0
                color #000000
            }
            element "SaaS Platform" {
                shape RoundedBox
            }
            element "REST API" {
                shape Hexagon
            }
            element "InProgress" {
                opacity 50
                background #ffcc99
                stroke #ff6600
                color #000000
            }
            element "Completed" {
                background #90EE90
                stroke #228B22
                strokeWidth 3
                color #000000
            }

            relationship "Relationship" {
                thickness 2
            }
        }

        themes default
    }
}
