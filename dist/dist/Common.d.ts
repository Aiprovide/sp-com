/// <reference types="sharepoint" />
import "Promise";
import "jquery-ui";
export declare enum BrowserKind {
    Other = 0,
    IE10Below = 1,
    IE11 = 2,
    Edge = 3,
    Chrome = 4,
    FireFox = 5,
    Safari = 6,
    Opera = 7,
}
export declare class Common {
    constructor();
    static find(target: any[], func: (value: any, index: number, array: any[]) => boolean): any;
    static WaitLoadPage(retFunc: (retVal: boolean) => void, value: boolean): void;
    static WaitSPJSLoaded(retFunc: () => void): void;
    static WaitSPUIJSLoaded(retFunc: () => void): void;
    static Get_SPLoginUser(spClientContext: SP.ClientContext, spWeb: SP.Web): Promise<SP.User>;
    static Get_SPGivenUser(spClientContext: SP.ClientContext, loginName: string, spWeb: SP.Web): Promise<SP.User>;
    static Get_SPGivenUsers(spClientContext: SP.ClientContext, loginNames: string[], spWeb: SP.Web): Promise<SP.User[]>;
    static Get_SPUserByID(spClientContext: SP.ClientContext, userId: number, spWeb: SP.Web): Promise<SP.User>;
    static Get_SPUsersFromFieldUsers(spClientContext: SP.ClientContext, fieldUsers: SP.FieldUserValue[], spWeb: SP.Web): Promise<SP.User[]>;
    static Get_SPUserFromFieldUser(spClientContext: SP.ClientContext, fieldUser: SP.FieldUserValue, spWeb: SP.Web): Promise<SP.User>;
    static IS_SPUserInFieldUsers(user: SP.User, spListItem: SP.ListItem, fieldName: string): boolean;
    static get_RealSPUsers(spClientContext: SP.ClientContext, users: SP.User[]): Promise<void>;
    static Get_SPListItemCollection(spClientContext: SP.ClientContext, spList: SP.List, query: string, folderurl?: string): Promise<SP.ListItem[]>;
    static Get_SPSubFolderCollection(spClientContext: SP.ClientContext, spFolder: SP.Folder): Promise<SP.Folder[]>;
    static Get_SPSubListItemCollection(spClientContext: SP.ClientContext, spFolder: SP.Folder, spList: SP.List): Promise<SP.ListItem[]>;
    static Get_SPListItemByID(spClientContext: SP.ClientContext, spList: SP.List, ID: number): SP.ListItem;
    static Get_MaxIDInSPFolder(spClientContext: SP.ClientContext, folder: SP.Folder, spList: SP.List): Promise<number>;
    static Get_MinIDInSPFolder(spClientContext: SP.ClientContext, folder: SP.Folder, spList: SP.List): Promise<number>;
    static Add_SPListFolder(spClientContext: SP.ClientContext, folder: SP.Folder, spList: SP.List, props: {
        key: string;
        value: any;
    }[]): Promise<SP.ListItem>;
    static Move_SPListFolder(spClientContext: SP.ClientContext, folder: SP.Folder, toPfolder: SP.Folder, serverUrl: string): Promise<void>;
    static Move_SPDocLibFolder(spClientContext: SP.ClientContext, folder: SP.Folder, toPfolder: SP.Folder, serverUrl: string): Promise<void>;
    static Add_SPListFolderItem(spClientContext: SP.ClientContext, folder: SP.Folder, spList: SP.List, props: {
        key: string;
        value: any;
    }[]): Promise<SP.ListItem>;
    static Copy_SPListItem(spClientContext: SP.ClientContext, spListItem: SP.ListItem, toPfolder: SP.Folder, serverUrl: string): Promise<void>;
    static Copy_SPDocLibItem(spClientContext: SP.ClientContext, spListItem: SP.ListItem, toPfolder: SP.Folder, serverUrl: string): Promise<void>;
    static Move_SPListItem(spClientContext: SP.ClientContext, spListItem: SP.ListItem, toPfolder: SP.Folder, serverUrl: string): Promise<void>;
    static Move_SPDocLibItem(spClientContext: SP.ClientContext, spListItem: SP.ListItem, toPfolder: SP.Folder, serverUrl: string): Promise<void>;
    static Delete_SPListFolder(spClientContext: SP.ClientContext, folder: SP.Folder, spList: SP.List, recycle?: boolean): Promise<void>;
    static Add_SPListItem(spClientContext: SP.ClientContext, spList: SP.List, props: {
        key: string;
        value: any;
    }[]): Promise<SP.ListItem>;
    static Edit_SPListItem(spClientContext: SP.ClientContext, spList: SP.List, spListItem: SP.ListItem, props: {
        key: string;
        value: any;
    }[]): Promise<void>;
    static Remove_SPListItem(spClientContext: SP.ClientContext, spList: SP.List, spListItem: SP.ListItem, recycle?: boolean): Promise<void>;
    static Edit_SPListItemByID(spClientContext: SP.ClientContext, spList: SP.List, ID: number, props: {
        key: string;
        value: any;
    }[]): Promise<void>;
    static Remove_SPListItemByID(spClientContext: SP.ClientContext, spList: SP.List, ID: number, recycle?: boolean): Promise<void>;
    static Get_SPListByTitle(spClientContext: SP.ClientContext, spWeb: SP.Web, title: string): Promise<SP.List>;
    static Get_SPListsByTitles(spClientContext: SP.ClientContext, spWeb: SP.Web, titles: string[]): Promise<SP.List[]>;
    static Get_SPListByURL(spClientContext: SP.ClientContext, spWeb: SP.Web, url: string): Promise<SP.List>;
    static Get_SPListByServerRelativeUrl(spClientContext: SP.ClientContext, spWeb: SP.Web, serverRelativeUrl: string): Promise<SP.List>;
    static Add_SPList(spClientContext: SP.ClientContext, spWeb: SP.Web, title: string, url: string, type: SP.ListTemplateType, descript?: string): Promise<SP.List>;
    static Add_SPFields(spClientContext: SP.ClientContext, spList: SP.List, spFields: {
        FieldIname: string;
        FieldDname: string;
        FieldType: SP.FieldType;
        FieldProp: string;
        IsIndex: boolean;
        DefValue: string;
    }[]): Promise<boolean>;
    static get_FileName(file: File): string;
    static ReadAsText(file: File): Promise<string>;
    static ReadAsBinary(file: File): Promise<any>;
    static UploadFileToDocLib(arrayBuffer: any, serverRelativeUrlToFolder: string, fileName: string, serverUrl: string, siteUrl: string, appUrl: string): Promise<void>;
    static Get_SPSearchExecutorResult(spClientContext: SP.ClientContext, keywordquery: string): Promise<SP.JsonObjectResult>;
    static SendEmail(spClientContext: SP.ClientContext, toUserList: SP.User[], subject: string, body: string, appUrl: string): Promise<void>;
    static sendMailReplacer(value: string): string;
    static CrToBrReplacer(value: string): string;
    static arrayBufferToBase64(buffer: any): string;
    static DoesUserHaveWebPermissions(spClientContext: SP.ClientContext, spWeb: SP.Web, permission: SP.PermissionKind): Promise<boolean>;
    static DoesUserHaveListPermissions(spClientContext: SP.ClientContext, spList: SP.List, permission: SP.PermissionKind): Promise<boolean>;
    static DoesSpecifyUserHaveListPermissions(spClientContext: SP.ClientContext, spList: SP.List, spUser: SP.User, permission: SP.PermissionKind): Promise<boolean>;
    static SPExecuteQueryAsync(spClientContext: SP.ClientContext, ...clientObjects: SP.ClientObject[]): Promise<void>;
    static SPExecuteArrayQueryAsync(spClientContext: SP.ClientContext, clientObjects: SP.ClientObject[]): Promise<void>;
    static SPExecutePropQueryAsync(spClientContext: SP.ClientContext, clientObject: SP.ClientObject, property: string): Promise<void>;
    static SPExecuteArryPropQueryAsync(spClientContext: SP.ClientContext, propArray: {
        clientObject: SP.ClientObject;
        property: string;
    }[]): Promise<void>;
    static Sleep(time: number): Promise<void>;
    static GetBrowserKind(): BrowserKind;
    static OpenWindow(strPage: string, strParam: string, strName: string): void;
    static OpenWindowWithOutSPHostUrl(strPage: string, strParam: string, strName: string): void;
    static OpenTabWindow(url: string): void;
    static OpenDocument(url: string): void;
    static get_ParameterByName(name: string): string;
    static get_ParameterString(): string;
    static get_MonthString(date: Date, language?: string): string;
    static get_DateString(date: Date, language?: string): string;
    static get_TimeString(date: Date): string;
    static get_FormatDate(date: Date, format?: string): string;
    static generateGUID(): string;
    static DoseUseOWA(fileName: string): boolean;
    static HtmlEscape(str: any): any;
}
