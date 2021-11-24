import { Components, ContextInfo, List, Types, Web } from "gd-sprest-bs";
import * as moment from "moment";
import Strings from "./strings";

// Event Item
export interface IEventItem extends Types.SP.ListItemOData {
    Description: string;
    StartDate: string;
    EndDate: string;
    EventStatus: { results: string };
    IsCancelled: boolean;
    Location: string;
    OpenSpots: string;
    Capacity: string;
    CreatedBy: string;
    POC: {
        results: {
            EMail: string;
            Id: number;
            Title: string;
        }[]
    };
    POCId: { results: number[] };
    RegisteredUsers: {
        results: {
            EMail: string;
            Id: number;
            Title: string;
        }[]
    };
    RegisteredUsersId: { results: number[] };
    WaitListedUsers: {
        results: {
            EMail: string;
            Id: number;
            Title: string;
        }[]
    };
    WaitListedUsersId: { results: number[] };
}

// Configuration
export interface IConfiguration {
    adminGroupName?: string;
    headerImage?: string;
    headerTitle?: string;
    hideAddToCalendarColumn?: boolean;
    hideHeader?: boolean;
    membersGroupName?: string;
}
/**
 * Data Source
 */
export class DataSource {
    // Filter Set
    private static _filterSet: boolean = false;
    static get FilterSet(): boolean { return this._filterSet; }
    static SetFilter(filterSet: boolean) {
        this._filterSet = filterSet;
    }

    // Events
    private static _events: IEventItem[] = null;
    static get Events(): IEventItem[] { return this._events; }
    // static get ActiveEvents(): IEventItem[] {
    //     let activeEvents: IEventItem[] = [];
    //     let today = moment();
    //     this._events.forEach((event) => {
    //         let startDate = event.StartDate;
    //         if(moment(startDate).isAfter(today)) {
    //             activeEvents.push(event);
    //         }
    //     })
    //     return activeEvents;
    // }

    static loadEvents(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            let web = Web();

            // Load the data
            if (this._isAdmin) {
                web.Lists(Strings.Lists.Events).Items().query({
                    Expand: ["AttachmentFiles", "POC", "RegisteredUsers", "WaitListedUsers"],
                    GetAllItems: true,
                    OrderBy: ["StartDate asc"],
                    Top: 5000,
                    Select: [
                        "*", "POC/Id", "POC/Title", "POC/EMail",
                        "RegisteredUsers/Id", "RegisteredUsers/Title", "RegisteredUsers/EMail",
                        "WaitListedUsers/Id", "WaitListedUsers/Title", "WaitListedUsers/EMail"
                    ]
                }).execute(
                    // Success
                    items => {
                        // Resolve the request
                        this._events = items.results as any;
                    },
                    // Error
                    () => { reject(); }
                );
            }
            else {
                let today = moment().toISOString();
                web.Lists(Strings.Lists.Events).Items().query({
                    Expand: ["AttachmentFiles", "POC", "RegisteredUsers", "WaitListedUsers"],
                    Filter: `StartDate ge '${today}'`,
                    GetAllItems: true,
                    OrderBy: ["StartDate asc"],
                    Top: 5000,
                    Select: [
                        "*", "POC/Id", "POC/Title",
                        "RegisteredUsers/Id", "RegisteredUsers/Title", "RegisteredUsers/EMail",
                        "WaitListedUsers/Id", "WaitListedUsers/Title", "WaitListedUsers/EMail"
                    ]
                }).execute(
                    items => {
                        // Resolve the request
                        this._events = items.results as any;
                    },
                    () => { reject(); }
                );
            }
            // Load the user permissions for the Events list
            web.Lists(Strings.Lists.Events).getUserEffectivePermissions(this._userLoginName).execute(perm => {
                // Save the user permissions
                this._eventRegPerms = perm.GetUserEffectivePermissions;
            }, () => {
                // Unable to determine the user permissions
                this._eventRegPerms = {};
            });

            // Once both queries are complete, return promise
            web.done(() => {
                // Resolve the request
                resolve();
            });
        });
    }

    // Required Docs
    private static _docs: IEventItem[] = null;
    static get Docs(): IEventItem[] { return this._docs; }
    static loadDocs(): PromiseLike<void> {
        return new Promise((resolve, reject) => {

            let web = Web();

            // Load the data
            web.Lists(Strings.Lists.Templates).RootFolder().query({
                Expand: ["Folders", "Folders/Files", "Folders/Files/Author", "Folders/Files/ListItemAllFields", "Folders/Files/ModifiedBy", "Files", "Files/Author", "Files/ListItemAllFields", "Files/ModifiedBy"],
                GetAllItems: true,
                Top: 5000,
                // Select: [
                //     "*", "Modified"
                // ]
            }).execute(folder => {
                // Set the folders and files
                this._files = folder.Files.results;
                this._folders = [];
                for (let i = 0; i < folder.Folders.results.length; i++) {
                    let subFolder = folder.Folders.results[i];
                    // Ignore the OTB Forms internal folder  
                    if (subFolder.Name != "Forms") { this._folders.push(subFolder as any); }
                }

                // Resolve the request
                resolve();
            }, reject);

            // Once both queries are complete, return promise
            web.done(() => {
                // Resolve the request
                resolve();
            });
        });
    }

    // Event Registration Permissions
    private static _eventRegPerms: Types.SP.BasePermissions;
    static get EventRegPerms(): Types.SP.BasePermissions { return this._eventRegPerms; };

    // Check if user is an admin
    private static _isAdmin: boolean = false;
    static get IsAdmin(): boolean { return this._isAdmin; }

    // Set Admin status
    private static GetAdminStatus(): PromiseLike<void> {
        return new Promise((resolve) => {
            if (this._cfg.adminGroupName) {
                Web().SiteGroups().getByName(this._cfg.adminGroupName).Users().getById(ContextInfo.userId).execute(
                    () => { this._isAdmin = true; resolve(); },
                    () => { this._isAdmin = false; resolve(); }
                )
            }
            else {
                Web().AssociatedOwnerGroup().Users().getById(ContextInfo.userId).execute(
                    () => { this._isAdmin = true; resolve(); },
                    () => { this._isAdmin = false; resolve(); }
                )
            }
        });
    }

    // Status Filters
    private static _statusFilters: Components.ICheckboxGroupItem[] = [{
        label: "Show inactive events",
        type: Components.CheckboxGroupTypes.Switch,
        isSelected: false
    }];
    static get StatusFilters(): Components.ICheckboxGroupItem[] { return this._statusFilters; }

    // User Login Name
    private static _userLoginName = null;
    static get UserLoginName(): string { return this._userLoginName; }
    private static loadUserName(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Get the user information
            Web().CurrentUser().execute(user => {
                // Set the login name
                this._userLoginName = user.LoginName;

                // Resolve the request
                resolve();
            }, reject);
        });
    }

    // Loads the list data
    static init(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Load the optional configuration file
            this.loadConfiguration().then(() => {
                // Load the user's login name
                this.loadUserName().then(() => {
                    // Load the security group urls
                    this.loadSecurityGroupUrls().then(() => {
                        // Determine if the user is an admin
                        this.GetAdminStatus().then(() => {
                            // Load the events
                            this.loadEvents().then(() => {
                                this.loadDocs().then(() => {
                                    // Resolve the request
                                    resolve();
                                }, reject);
                            }, reject);
                        }, reject);
                    }, reject);
                }, reject);
            }, reject);
        });
    }

    // Configuration
    private static _cfg: IConfiguration = null;
    static get Configuration(): IConfiguration { return this._cfg; }
    static loadConfiguration(): PromiseLike<void> {
        // Return a promise
        return new Promise(resolve => {
            // Get the current web
            Web().getFileByServerRelativeUrl(Strings.EventRegConfig).content().execute(
                // Success
                file => {
                    // Convert the string to a json object
                    let cfg = null;
                    try { cfg = JSON.parse(String.fromCharCode.apply(null, new Uint8Array(file))); }
                    catch { cfg = {}; }

                    // Set the configuration
                    this._cfg = cfg;

                    // Resolve the request
                    resolve();
                },

                // Error
                () => {
                    // Set the configuration to nothing
                    this._cfg = {} as any;

                    // Resolve the request
                    resolve();
                }
            );
        });
    }

    // Templates Files 
    private static _files: Types.SP.File[];
    static get Files(): Types.SP.File[] { return this._files; }
    private static _folders: Types.SP.FolderOData[];
    static get Folders(): Types.SP.FolderOData[] { return this._folders; }
    static loadTemplateFiles(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Load the library
            List(Strings.Lists.Templates).RootFolder().query({
                Expand: [
                    "Folders", "Folders/Files", "Folders/Files/Author", "Folders/Files/ListItemAllFields", "Folders/Files/ModifiedBy",
                    "Files", "Files/Author", "Files/ListItemAllFields", "Files/ModifiedBy", "Files/CreatedBy"
                ]
            }).execute(folder => {
                // Set the folders and files
                this._files = folder.Files.results;
                this._folders = [];
                for (let i = 0; i < folder.Folders.results.length; i++) {
                    let subFolder = folder.Folders.results[i];
                    // Ignore the OTB Forms internal folder  
                    if (subFolder.Name != "Forms") { this._folders.push(subFolder as any); }
                }
                // Resolve the request
                resolve();
            }, reject);
        });
    }

    // Security Groups
    private static _managerId: number = null;
    static get ManagersUrl(): string { return ContextInfo.webServerRelativeUrl + "/_layouts/15/people.aspx?MembershipGroupId=" + this._managerId; }
    private static _memberId: number = null;
    static get MembersUrl(): string { return ContextInfo.webServerRelativeUrl + "/_layouts/15/people.aspx?MembershipGroupId=" + this._memberId; }
    static get DocLibUrl(): string { return ContextInfo.webServerRelativeUrl + "/RequiredDocs/Forms/AllItems.aspx"; }
    static loadSecurityGroupUrls(): PromiseLike<void> {
        return new Promise((resolve) => {
            let web = Web();

            // Load the owner's group
            let ownersGroup = DataSource.Configuration.adminGroupName;
            (ownersGroup ? Web().SiteGroups().getByName(ownersGroup) : Web().AssociatedOwnerGroup()).execute(group => {
                // Set the id
                this._managerId = group.Id;
            });

            // Load the member's group
            let membersGroup = DataSource.Configuration.membersGroupName;
            (membersGroup ? Web().SiteGroups().getByName(membersGroup) : Web().AssociatedMemberGroup()).execute(group => {
                // Set the id
                this._memberId = group.Id;
            });

            // Wait for the requests to complete
            web.done(() => { resolve(); });
        });
    }
}