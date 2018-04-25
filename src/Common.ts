import "Promise";
//import $ from 'jquery';
import "jquery-ui";

// ブラウザのバージョン
export enum BrowserKind
{
    Other = 0,
    IE10Below = 1,
    IE11 = 2,
    Edge = 3,
    Chrome = 4,
    FireFox = 5,
    Safari = 6,
    Opera = 7
}

export class Common
{
    //
    // コンストラクタ
    //
    constructor()
    {
    }

    //
    // .find()の実装
    //
    public static find(target: any[], func: (value: any, index: number, array: any[]) => boolean): any
    {
        let fary: any[] = target.filter(func);
        let fobj: any = null;
        if (fary.length > 0)
        {
            fobj = fary[0];
        }

        return fobj;
    }

    //
    // ページのロード完了を待つ
    //
    public static WaitLoadPage(retFunc: (retVal: boolean) => void, value: boolean): void
    {
        $((): void =>
        {
            retFunc(value);
        });
    }


    //
    // SharePointライブラリ sp.js が読み込まれるまで待機する
    //
    public static WaitSPJSLoaded(retFunc: () => void): void
    {
        if (SP.SOD !== null && typeof SP.SOD !== "undefined")
        {
            SP.SOD.executeFunc('sp.js', 'SP.ClientContext', function ()
            {
                console.log("Initiating SP.ClientContext");
                SP.SOD.executeOrDelayUntilScriptLoaded(retFunc, "sp.js");
            });
        }
        else
        {
            retFunc();
        }
    }

    //
    // SharePointライブラリ sp.jsとsp.ui.controls.js が読み込まれるまで待機する
    //
    public static WaitSPUIJSLoaded(retFunc: () => void): void
    {
        $((): void =>
        {
            if (SP.SOD !== null && typeof SP.SOD !== "undefined")
            {
                SP.SOD.executeFunc('sp.js', 'SP.ClientContext', (): void =>
                {
                    console.log("Initiating SP.ClientContext");
                    SP.SOD.executeOrDelayUntilScriptLoaded((): void =>
                    {
                        SP.SOD.executeOrDelayUntilScriptLoaded((): void =>
                        {
                            retFunc();
                        }, "sp.ui.controls.js");
                    }, "sp.js");
                });
            }
            else
            {
                retFunc();
            }
        });
    }

    //
    // 現在のログインユーザーを取得してインスタンス化して返す
    //
    public static async Get_SPLoginUser(spClientContext: SP.ClientContext, spWeb: SP.Web): Promise<SP.User>
    {
        // 現在のログインユーザーを取得してインスタンス化する
        let loginUser: SP.User = spWeb.get_currentUser();
        await Common.SPExecuteQueryAsync(spClientContext, loginUser);
        return loginUser;
    }

    //
    // 与えられたログイン名のユーザーを取得してインスタンス化して返す
    //
    public static async Get_SPGivenUser(spClientContext: SP.ClientContext, loginName: string, spWeb: SP.Web): Promise<SP.User>
    {
        // 与えられたログイン名のユーザーを取得してインスタンス化する
        let user: SP.User = spWeb.ensureUser(loginName);
        await Common.SPExecuteQueryAsync(spClientContext, user);
        return user;
    }

    //
    // 与えられたログイン名のユーザー(複数)を取得してインスタンス化して返す
    //
    public static async Get_SPGivenUsers(spClientContext: SP.ClientContext, loginNames: string[], spWeb: SP.Web): Promise<SP.User[]>
    {
        // 与えられたログイン名のユーザーを取得してインスタンス化する
        let users: SP.User[] = new Array<SP.User>();
        loginNames.forEach((value: string, index: number, array: string[]): void =>
        {
            let user: SP.User = spWeb.ensureUser(value);
            users.push(user);
        });
        await Common.SPExecuteArrayQueryAsync(spClientContext, users);
        return users;
    }

    //
    // 与えられたユーザーIDのユーザーを取得してインスタンス化して返す
    //
    public static async Get_SPUserByID(spClientContext: SP.ClientContext, userId: number, spWeb: SP.Web): Promise<SP.User>
    {
        // 与えられたユーザーIDのユーザーを取得してインスタンス化する
        let user: SP.User = spWeb.get_siteUsers().getById(userId);
        await Common.SPExecuteQueryAsync(spClientContext, user);
        return user;
    }

    //
    // 与えられたSP.FieldUserValue[]に設定されたユーザーを、SP.User[]にする
    //
    public static async Get_SPUsersFromFieldUsers(spClientContext: SP.ClientContext, fieldUsers: SP.FieldUserValue[], spWeb: SP.Web): Promise<SP.User[]>
    {
        let spUsers: SP.User[] = new Array<SP.User>();
        fieldUsers.forEach((value: SP.FieldUserValue, index: number, array: SP.FieldUserValue[]): void =>
        {
            var userID: number = value.get_lookupId();
            spUsers.push(spWeb.get_siteUsers().getById(userID));
        });

        // 与えられたユーザーの配列の各要素のユーザーを実体化する
        await Common.get_RealSPUsers(spClientContext, spUsers);
        return spUsers;
    }

    //
    // 与えられたSP.FieldUserValueに設定されたユーザーを、SP.Userにする
    //
    public static async Get_SPUserFromFieldUser(spClientContext: SP.ClientContext, fieldUser: SP.FieldUserValue, spWeb: SP.Web): Promise<SP.User>
    {
        let spUsers: SP.User[] = new Array<SP.User>();
        var userID: number = fieldUser.get_lookupId();
        spUsers.push(spWeb.get_siteUsers().getById(userID));

        // ユーザーを実体化する
        await Common.get_RealSPUsers(spClientContext, spUsers);
        return spUsers[0];
    }

    //
    // 与えられたリストアイテムの与えられたSP.FieldUserValue[]列に、与えられたユーザーが存在するかどうかを返す
    // ユーザーとリストアイテムは実体化されていること
    //
    public static IS_SPUserInFieldUsers(user: SP.User, spListItem: SP.ListItem, fieldName: string)
    {
        // 与えられたユーザーが、伝言の通知先に存在するか
        let toUserValue: SP.FieldUserValue = spListItem.get_item(fieldName);
        let toUserValueID: number = toUserValue.get_lookupId();

        let userID: number = user.get_id();
        if (toUserValueID === userID)
        {
            return true;
        }
        else
        {
            return false;
        }
    }

    //
    // 与えられたユーザーの配列の各要素のユーザーを実体化する
    //
    public static async get_RealSPUsers(spClientContext: SP.ClientContext, users: SP.User[]): Promise<void>
    {
        let arglArray: SP.ClientObject[] = new Array<SP.ClientObject>();
        users.forEach((value: SP.User, index: number, array: SP.User[]): void =>
        {
            // serverObjectIsNull == false のものがリストアイテムとして実体化が可能
            if (value.get_serverObjectIsNull() === false)
            {
                arglArray.push(value);
            }
            else
            {
                value = null;
            }
        });
        await Common.SPExecuteArrayQueryAsync(spClientContext, arglArray);
        return;
    }

    //
    // 与えられたリストから、与えられたクエリーで検索したリストアイテムを返す
    //
    public static async Get_SPListItemCollection(spClientContext: SP.ClientContext, spList: SP.List, query: string, folderurl: string = ""): Promise<SP.ListItem[]>
    {
        let colSPListItem: SP.ListItem[] = new Array<SP.ListItem>();

        // クエリーの作成
        let objSPQuery: SP.CamlQuery = new SP.CamlQuery();
        objSPQuery.set_viewXml(query);
        if (folderurl !== null && folderurl.length > 0)
        {
            objSPQuery.set_folderServerRelativeUrl(folderurl);
        }

        let spListItems: SP.ListItemCollection = spList.getItems(objSPQuery);

        try
        {
            await Common.SPExecuteQueryAsync(spClientContext, spListItems);
            // 全リストアイテムを取得
            let itemEnumerator: IEnumerator<SP.ListItem> = spListItems.getEnumerator();
            while (itemEnumerator.moveNext())
            {
                let spListItem: SP.ListItem = itemEnumerator.get_current();
                colSPListItem.push(spListItem);
            }
            return colSPListItem;
        }
        catch (e)
        {
            return null;
        }
    }

    //
    // 与えられたフォルダの配下のサブフォルダを取得して返す
    //
    public static async Get_SPSubFolderCollection(spClientContext: SP.ClientContext, spFolder: SP.Folder): Promise<SP.Folder[]>
    {
        let colSPFolder: SP.Folder[] = new Array<SP.Folder>();

        // サブフォルダの取得
        let subspFolders: SP.FolderCollection = spFolder.get_folders();

        try
        {
            await Common.SPExecuteQueryAsync(spClientContext, subspFolders);
            // 全リストアイテムを取得
            let itemEnumerator: IEnumerator<SP.Folder> = subspFolders.getEnumerator();
            while (itemEnumerator.moveNext())
            {
                let folder: SP.Folder = itemEnumerator.get_current();
                colSPFolder.push(folder);
            }
            return colSPFolder;
        }
        catch (e)
        {
            return null;
        }
    }

    //
    // 与えられたフォルダの配下のリストアイテムを取得して返す
    //
    public static async Get_SPSubListItemCollection(spClientContext: SP.ClientContext, spFolder: SP.Folder, spList: SP.List): Promise<SP.ListItem[]>
    {
        let camlQuery: string = "<View><Query><Where></Where>"
            + "<OrderBy><FieldRef Name='ID' Ascending='TRUE'></FieldRef></OrderBy>"
            + "</Query></View>";
        let folderUrl = spFolder.get_serverRelativeUrl();

        // 与えられたリストから、与えられたクエリーで検索したリストアイテムを返す
        let colListItem: SP.ListItem[] = await Common.Get_SPListItemCollection(spClientContext, spList, camlQuery, folderUrl);

        return colListItem;
    }

    //
    // リストアイテムIDで指定されたリストアイテムを取得して返す
    //
    public static Get_SPListItemByID(spClientContext: SP.ClientContext, spList: SP.List, ID: number): SP.ListItem
    {
        let spListItem: SP.ListItem = spList.getItemById(ID);
        return spListItem;

        //let camlQuery: string = "<View><Query><Where><Eq><FieldRef Name='ID' /><Value Type='Number'>" + ID + "</Value></Eq></Where></Query><RowLimit>10</RowLimit></View>";

        //// 与えられたリストから、与えられたIDで検索したリストアイテムを返す
        //let colListItem: SP.ListItem[] = await Common.Get_SPListItemCollection(spClientContext, spList, camlQuery);

        //if (colListItem !== null && colListItem.length > 0)
        //{
        //    return colListItem[0];
        //}
        //else
        //{
        //    return null;
        //}
    }

    //
    // 与えられたフォルダのIDの最大値を取得する
    //
    public static async Get_MaxIDInSPFolder(spClientContext: SP.ClientContext, folder: SP.Folder, spList: SP.List): Promise<number>
    {
        let camlQuery: string = "<View><Query><Where>"
            + "<Eq><FieldRef Name='FSObjType'/><Value Type='Integer'>0</Value></Eq>"
            + "</Where>"
            + "<OrderBy><FieldRef Name='ID' Ascending='FALSE'></FieldRef></OrderBy>"
            + "</Query><RowLimit>1</RowLimit></View>";
        let folderUrl = folder.get_serverRelativeUrl();

        // 与えられたリストから、与えられたクエリーで検索したリストアイテムを返す
        let colListItem: SP.ListItem[] = await Common.Get_SPListItemCollection(spClientContext, spList, camlQuery, folderUrl);

        if (colListItem.length > 0)
        {
            let maxID: number = colListItem[0].get_id();
            return maxID;
        }
        else
        {
            return 0;
        }
    }

    //
    // 与えられたフォルダのIDの最小値を取得する
    //
    public static async Get_MinIDInSPFolder(spClientContext: SP.ClientContext, folder: SP.Folder, spList: SP.List): Promise<number>
    {
        let camlQuery: string = "<View><Query><Where>"
            + "<Eq><FieldRef Name='FSObjType'/><Value Type='Integer'>0</Value></Eq>"
            + "</Where>"
            + "<OrderBy><FieldRef Name='ID' Ascending='TRUE'></FieldRef></OrderBy>"
            + "</Query><RowLimit>1</RowLimit></View>";
        let folderUrl = folder.get_serverRelativeUrl();

        // 与えられたリストから、与えられたクエリーで検索したリストアイテムを返す
        let colListItem: SP.ListItem[] = await Common.Get_SPListItemCollection(spClientContext, spList, camlQuery, folderUrl);

        if (colListItem.length > 0)
        {
            let minID: number = colListItem[0].get_id();
            return minID;
        }
        else
        {
            return 0;
        }
    }

    //
    // リストの与えられたフォルダに、サブフォルダを追加する
    //
    public static async Add_SPListFolder(spClientContext: SP.ClientContext, folder: SP.Folder, spList: SP.List, props: { key: string; value: any }[]): Promise<SP.ListItem>
    {
        // リストアイテムを作成する
        let listItemCreateInfo: SP.ListItemCreationInformation = new SP.ListItemCreationInformation();
        listItemCreateInfo.set_underlyingObjectType(SP.FileSystemObjectType.folder);
        listItemCreateInfo.set_leafName(this.generateGUID());
        listItemCreateInfo.set_folderUrl(folder.get_serverRelativeUrl());

        let spListItem: SP.ListItem = spList.addItem(listItemCreateInfo);

        props.forEach(
            function (property, index, array)
            {
                spListItem.set_item(property.key, property.value);
            });

        spListItem.update();
        await Common.SPExecuteQueryAsync(spClientContext, spListItem);

        // フォルダ名を変更する
        let id: number = spListItem.get_id();
        spListItem.set_item("FileLeafRef", id.toString());
        spListItem.update();
        await Common.SPExecuteQueryAsync(spClientContext, spListItem);

        return spListItem;
    }

    //
    // リストの与えられたフォルダを、別のフォルダの配下に移動する
    //
    public static async Move_SPListFolder(spClientContext: SP.ClientContext, folder: SP.Folder, toPfolder: SP.Folder, serverUrl: string): Promise<void>
    {
        let srcsrUrl: string = folder.get_serverRelativeUrl();
        let sfolderName: string = folder.get_name();
        let dstsrUrl: string = toPfolder.get_serverRelativeUrl();
        let srcUrl: string = serverUrl + srcsrUrl;
        let dstUrl: string = serverUrl + dstsrUrl;
        dstUrl = (dstUrl.endsWith("/")) ? dstUrl + sfolderName : dstUrl + "/" + sfolderName;
        SP.MoveCopyUtil.moveFolder(spClientContext, srcUrl, dstUrl);

        // フォルダを移動する
        await Common.SPExecuteQueryAsync(spClientContext);

        return;
    }

    //
    // ドキュメントライブラリの与えられたフォルダを、別のフォルダの配下に移動する
    //
    public static async Move_SPDocLibFolder(spClientContext: SP.ClientContext, folder: SP.Folder, toPfolder: SP.Folder, serverUrl: string): Promise<void>
    {
        let srcsrUrl: string = folder.get_serverRelativeUrl();
        let sfolderName: string = folder.get_name();
        let dstsrUrl: string = toPfolder.get_serverRelativeUrl();
        let srcUrl: string = serverUrl + srcsrUrl;
        let dstUrl: string = serverUrl + dstsrUrl;
        dstUrl = (dstUrl.endsWith("/")) ? dstUrl + sfolderName : dstUrl + "/" + sfolderName;
        SP.MoveCopyUtil.moveFolder(spClientContext, srcUrl, dstUrl);

        // フォルダを移動する
        await Common.SPExecuteQueryAsync(spClientContext);

        return;
    }

    //
    // リストの与えられたフォルダに、アイテムを追加する
    //
    public static async Add_SPListFolderItem(spClientContext: SP.ClientContext, folder: SP.Folder, spList: SP.List, props: { key: string; value: any }[]): Promise<SP.ListItem>
    {
        // リストアイテムを作成する
        let listItemCreateInfo: SP.ListItemCreationInformation = new SP.ListItemCreationInformation();
        listItemCreateInfo.set_underlyingObjectType(SP.FileSystemObjectType.file);
        listItemCreateInfo.set_folderUrl(folder.get_serverRelativeUrl());

        let spListItem: SP.ListItem = spList.addItem(listItemCreateInfo);

        props.forEach(
            function (property, index, array)
            {
                spListItem.set_item(property.key, property.value);
            });

        spListItem.update();

        // リストアイテムを追加する
        await Common.SPExecuteQueryAsync(spClientContext, spListItem);

        return spListItem;
    }

    //
    // リストの与えられたアイテムを、別のフォルダの配下にコピーする
    //
    public static async Copy_SPListItem(spClientContext: SP.ClientContext, spListItem: SP.ListItem, toPfolder: SP.Folder, serverUrl: string): Promise<void>
    {
        //
        // 以下は上手くいかない
        //

        //// SP.Fileを取得する
        //let argPropArray: { clientObject: SP.ClientObject; property: string }[] = new Array<{ clientObject: SP.ClientObject; property: string }>();
        //argPropArray.push({ clientObject: spListItem, property: "FileRef" });
        //argPropArray.push({ clientObject: spListItem, property: "FileLeafRef" });
        //await Common.SPExecuteArryPropQueryAsync(spClientContext, argPropArray);

        //let dstsrUrl: string = toPfolder.get_serverRelativeUrl();

        //let sfileName: string = spListItem.get_item('FileLeafRef');
        //let srcsrUrl: string = spListItem.get_item('FileRef');
        //let srcUrl: string = serverUrl + srcsrUrl;
        //let dstUrl: string = serverUrl + dstsrUrl;
        //dstUrl = (dstUrl.endsWith("/")) ? dstUrl + sfileName : dstUrl + "/" + sfileName;
        //SP.MoveCopyUtil.copyFile(spClientContext, srcUrl, dstUrl, true);

        //// ファイルをコピーする
        //await Common.SPExecuteQueryAsync(spClientContext);

        //return;
    }

    //
    // ドキュメントライブラリの与えられたアイテムを、別のフォルダの配下にコピーする
    //
    public static async Copy_SPDocLibItem(spClientContext: SP.ClientContext, spListItem: SP.ListItem, toPfolder: SP.Folder, serverUrl: string): Promise<void>
    {
        // SP.Fileを取得する
        let argPropArray: { clientObject: SP.ClientObject; property: string }[] = new Array<{ clientObject: SP.ClientObject; property: string }>();
        argPropArray.push({ clientObject: spListItem, property: "FileRef" });
        argPropArray.push({ clientObject: spListItem, property: "FileLeafRef" });
        await Common.SPExecuteArryPropQueryAsync(spClientContext, argPropArray);

        let dstsrUrl: string = toPfolder.get_serverRelativeUrl();

        let sfileName: string = spListItem.get_item('FileLeafRef');
        let srcsrUrl: string = spListItem.get_item('FileRef');
        let srcUrl: string = serverUrl + srcsrUrl;
        let dstUrl: string = serverUrl + dstsrUrl;
        dstUrl = (dstUrl.endsWith("/")) ? dstUrl + sfileName : dstUrl + "/" + sfileName;
        SP.MoveCopyUtil.copyFile(spClientContext, srcUrl, dstUrl, true);

        // ファイルをコピーする
        await Common.SPExecuteQueryAsync(spClientContext);

        return;
    }

    //
    // リストの与えられたアイテムを、別のフォルダの配下に移動する
    //
    public static async Move_SPListItem(spClientContext: SP.ClientContext, spListItem: SP.ListItem, toPfolder: SP.Folder, serverUrl: string): Promise<void>
    {
        // SP.Fileを取得する
        let argPropArray: { clientObject: SP.ClientObject; property: string }[] = new Array<{ clientObject: SP.ClientObject; property: string }>();
        argPropArray.push({ clientObject: spListItem, property: "FileRef" });
        argPropArray.push({ clientObject: spListItem, property: "FileLeafRef" });
        await Common.SPExecuteArryPropQueryAsync(spClientContext, argPropArray);

        //let spFile: SP.File = spListItem.get_file();
        //var fileUrl: string = spListItem.get_item('FileRef');
        let dstsrUrl: string = toPfolder.get_serverRelativeUrl();

        //let spFile: SP.File = spClientContext.get_web().getFileByServerRelativeUrl(fileUrl);
        //var targetfileUrl: string = fileUrl.replace(spListItem.get_item('FileDirRef'), dstsrUrl);

        //spFile.moveTo(targetfileUrl, SP.MoveOperations.overwrite);

        //let spFile: SP.File = spListItem.get_file();
        let sfileName: string = spListItem.get_item('FileLeafRef');
        let srcsrUrl: string = spListItem.get_item('FileRef');
        let srcUrl: string = serverUrl + srcsrUrl;
        let dstUrl: string = serverUrl + dstsrUrl;
        dstUrl = (dstUrl.endsWith("/")) ? dstUrl + sfileName : dstUrl + "/" + sfileName;
        SP.MoveCopyUtil.moveFile(spClientContext, srcUrl, dstUrl, true);

        // ファイルを移動する
        await Common.SPExecuteQueryAsync(spClientContext);

        return;
    }

    //
    // ドキュメントライブラリの与えられたアイテムを、別のフォルダの配下に移動する
    //
    public static async Move_SPDocLibItem(spClientContext: SP.ClientContext, spListItem: SP.ListItem, toPfolder: SP.Folder, serverUrl: string): Promise<void>
    {
        // SP.Fileを取得する
        let argPropArray: { clientObject: SP.ClientObject; property: string }[] = new Array<{ clientObject: SP.ClientObject; property: string }>();
        argPropArray.push({ clientObject: spListItem, property: "FileRef" });
        argPropArray.push({ clientObject: spListItem, property: "FileLeafRef" });
        await Common.SPExecuteArryPropQueryAsync(spClientContext, argPropArray);

        let dstsrUrl: string = toPfolder.get_serverRelativeUrl();

        let sfileName: string = spListItem.get_item('FileLeafRef');
        let srcsrUrl: string = spListItem.get_item('FileRef');
        let srcUrl: string = serverUrl + srcsrUrl;
        let dstUrl: string = serverUrl + dstsrUrl;
        dstUrl = (dstUrl.endsWith("/")) ? dstUrl + sfileName : dstUrl + "/" + sfileName;
        SP.MoveCopyUtil.moveFile(spClientContext, srcUrl, dstUrl, true);

        // ファイルを移動する
        await Common.SPExecuteQueryAsync(spClientContext);

        return;
    }

    //
    // リストの与えられたフォルダを削除する
    //
    public static async Delete_SPListFolder(spClientContext: SP.ClientContext, folder: SP.Folder, spList: SP.List, recycle: boolean = true): Promise<void>
    {
        if (recycle == true)
        {
            folder.recycle();
        }
        else
        {
            folder.deleteObject();
        }

        // フォルダを削除する
        await Common.SPExecuteQueryAsync(spClientContext, spList);

        return;
    }

    //
    // リストアイテムを追加する
    //
    public static async Add_SPListItem(spClientContext: SP.ClientContext, spList: SP.List, props: { key: string; value: any }[]): Promise<SP.ListItem>
    {
        // リストアイテムを作成する
        let listItemCreateInfo: SP.ListItemCreationInformation = new SP.ListItemCreationInformation();
        let spListItem: SP.ListItem = spList.addItem(listItemCreateInfo);

        props.forEach(
            function (property, index, array)
            {
                spListItem.set_item(property.key, property.value);
            });

        spListItem.update();

        // リストアイテムを追加する
        await Common.SPExecuteQueryAsync(spClientContext, spListItem);

        return spListItem;
    }

    //
    // 与えられたリストアイテムを更新する
    //
    public static async Edit_SPListItem(spClientContext: SP.ClientContext, spList: SP.List, spListItem: SP.ListItem, props: { key: string; value: any }[]): Promise<void>
    {
        props.forEach(
            function (property, index, array)
            {
                spListItem.set_item(property.key, property.value);
            });

        spListItem.update();
        // リストアイテムを更新する
        await Common.SPExecuteQueryAsync(spClientContext, spListItem);

        spList.update();
        await Common.SPExecuteQueryAsync(spClientContext, spList);
        return;

        //let itemID = spListItem.get_id();
        //// リストアイテムIDで指定されたリストアイテムを取得して返す
        //let updatedListItem: SP.ListItem = await Common.Get_SPListItemByID(spClientContext, spList, itemID);
        //// リストアイテムを実体化する
        //await Common.SPExecuteQueryAsync(spClientContext, updatedListItem);

        //return updatedListItem;
    }

    //
    // 与えられたリストアイテムを削除する
    //
    public static async Remove_SPListItem(spClientContext: SP.ClientContext, spList: SP.List, spListItem: SP.ListItem, recycle: boolean = true): Promise<void>
    {
        if (recycle == true)
        {
            spListItem.recycle();
        }
        else
        {
            spListItem.deleteObject();
        }

        // リストアイテムを更新する
        await Common.SPExecuteQueryAsync(spClientContext, spList);
    }

    //
    // IDで与えられたリストアイテムを更新する
    //
    public static async Edit_SPListItemByID(spClientContext: SP.ClientContext, spList: SP.List, ID: number, props: { key: string; value: any }[]): Promise<void>
    {
        // リストアイテムIDで指定されたリストアイテムを取得して返す
        let spListItem: SP.ListItem = Common.Get_SPListItemByID(spClientContext, spList, ID);

        props.forEach(
            function (property, index, array)
            {
                spListItem.set_item(property.key, property.value);
            });

        spListItem.update();

        // リストアイテムを更新する
        await Common.SPExecuteQueryAsync(spClientContext, spListItem);
    }

    //
    // IDで与えられたリストアイテムを削除する
    //
    public static async Remove_SPListItemByID(spClientContext: SP.ClientContext, spList: SP.List, ID: number, recycle: boolean = true): Promise<void>
    {
        // リストアイテムIDで指定されたリストアイテムを取得して返す
        let spListItem: SP.ListItem = Common.Get_SPListItemByID(spClientContext, spList, ID);

        if (recycle == true)
        {
            spListItem.recycle();
        }
        else
        {
            spListItem.deleteObject();
        }

        // リストアイテムを更新する
        await Common.SPExecuteQueryAsync(spClientContext, spListItem);
    }

    //
    // Webサイトの全リストから、与えられたタイトルのリストを取得して返す
    //
    public static async Get_SPListByTitle(spClientContext: SP.ClientContext, spWeb: SP.Web, title: string): Promise<SP.List>
    {
        let spList: SP.List = spWeb.get_lists().getByTitle(title);

        try
        {
            await Common.SPExecuteQueryAsync(spClientContext, spList);
            return spList;
        }
        catch (e)
        {
            return null;
        }
    }

    //
    // Webサイトの全リストから、与えられたタイトルのリストを取得して返す
    //
    public static async Get_SPListsByTitles(spClientContext: SP.ClientContext, spWeb: SP.Web, titles: string[]): Promise<SP.List[]>
    {
        let retSPList: SP.List[] = new Array<SP.List>();
        let colSPList: SP.List[] = new Array<SP.List>();
        titles.forEach ((value: string, index:number, array: string[]): void =>
        {
            let objSPList: SP.List = spWeb.get_lists().getByTitle(value);
            colSPList.push(objSPList);
            retSPList[value] = objSPList;
        });

        try
        {
            await Common.SPExecuteArrayQueryAsync(spClientContext, colSPList);
            return retSPList;
        }
        catch (e)
        {
            return null;
        }
    }

    //
    // Webサイトの全リストから、与えられたList URL(List/...)のリストを取得して返す
    //
    public static async Get_SPListByURL(spClientContext: SP.ClientContext, spWeb: SP.Web, url: string): Promise<SP.List>
    {
        let serverRelativeUrl: string = spWeb.get_serverRelativeUrl();
        if (serverRelativeUrl === "/")
        {
            serverRelativeUrl = "";
        }

        let fullUrl: string = serverRelativeUrl + "/" + url;
        let spList: SP.List = await this.Get_SPListByServerRelativeUrl(spClientContext, spWeb, fullUrl);

        return spList;
    }

    //
    // Webサイトの全リストから、与えられたServerRelative URLのリストを取得する
    //
    public static async Get_SPListByServerRelativeUrl(spClientContext: SP.ClientContext, spWeb: SP.Web, serverRelativeUrl: string): Promise<SP.List>
    {
        // 与えられたWebサイトに存在するリストの一覧を取得する
        let colSPLists: SP.ListCollection = spWeb.get_lists();

        await Common.SPExecuteQueryAsync(spClientContext, colSPLists);
        let arySPList: SP.List[] = new Array<SP.List>();

        // 全リストの一覧を取得
        let listEnumerator: IEnumerator<SP.List> = colSPLists.getEnumerator();
        while (listEnumerator.moveNext())
        {
            let objSPList: SP.List = listEnumerator.get_current();

            if (objSPList.get_hidden() === true)
            {
                // 隠しリストは処理しない
                continue;
            }
            arySPList.push(objSPList);
        }

        // 取得した一覧からURLが一致するリストを検索する
        for (let objSPList of arySPList)
        {
            let objRootFolder: SP.Folder = objSPList.get_rootFolder();
            await Common.SPExecuteQueryAsync(spClientContext, objSPList, objRootFolder);
            // 取得されたリストのURLを取得する
            let strURL: string = objRootFolder.get_serverRelativeUrl();

            if (serverRelativeUrl.toLowerCase() === strURL.toLowerCase())
            {
                // URLが一致するリストが見つかった
                return objSPList;
            }
        }

        return null;
    }

    //
    // リストを追加する
    //
    public static async Add_SPList(spClientContext: SP.ClientContext, spWeb: SP.Web, title: string, url: string, type: SP.ListTemplateType, descript: string = null): Promise<SP.List>
    {
        try
        {
            // リストを作成する
            let listCreateInfo: SP.ListCreationInformation = new SP.ListCreationInformation();
            listCreateInfo.set_title(title);
            listCreateInfo.set_url(url);
            listCreateInfo.set_templateType(type);
            if (descript !== null)
            {
                listCreateInfo.set_description(descript);
            }
            let colList: SP.ListCollection = spWeb.get_lists();
            await Common.SPExecuteQueryAsync(spClientContext, colList);

            // リストを追加する
            let newList: SP.List = colList.add(listCreateInfo);
            await Common.SPExecuteQueryAsync(spClientContext, newList);

            // 作成したリストを取得する
            let madeList: SP.List = spWeb.get_lists().getByTitle(title);
            await Common.SPExecuteQueryAsync(spClientContext, madeList);

            return madeList;
        }
        catch (e)
        {
            return null;
        }
    }

    //
    // 与えられたリストに、列を追加する
    //
    public static async Add_SPFields(spClientContext: SP.ClientContext, spList: SP.List, spFields: { FieldIname: string; FieldDname: string; FieldType: SP.FieldType; FieldProp: string; IsIndex: boolean; DefValue: string }[]): Promise<boolean>
    {
        try
        {
            // フィールドを作成する
            let fieldScm: string = "";

            // インデックスを付ける列
            let colIndexFields: SP.Field[] = new Array<SP.Field>();

            spFields.forEach((value: { FieldIname: string; FieldDname: string; FieldType: SP.FieldType; FieldProp: string; IsIndex: boolean; DefValue: string }, index: number, array: { FieldIname: string; FieldDname: string; FieldType: SP.FieldType; FieldProp: string; IsIndex: boolean; DefValue: string }[]): void =>
            {
                switch (value.FieldType)
                {
                    case SP.FieldType.text:
                        fieldScm = "<Field Type='Text'";
                        break;
                    case SP.FieldType.note:
                        fieldScm = "<Field Type='Note'";
                        break;
                    case SP.FieldType.number:
                        fieldScm = "<Field Type='Number'";
                        break;
                    case SP.FieldType.dateTime:
                        fieldScm = "<Field Type='DateTime'";
                        break;
                    case SP.FieldType.boolean:
                        fieldScm = "<Field Type='Boolean'";
                        break;
                    case SP.FieldType.user:
                        fieldScm = "<Field Type='User'";
                        break;
                    default:
                        break;
                }
                let fieldScm2: string = "";
                if (value.DefValue !== null)
                {
                    fieldScm2 = 
                        " Name = '"
                        + value.FieldIname
                        + "' StaticName='"
                        + value.FieldIname
                        + "' DisplayName='"
                        + value.FieldDname
                        + "' "
                        + value.FieldProp
                        + ">"
                        + "<Default>" + value.DefValue + "</Default>"
                        + "</Field>";
                }
                else
                {
                    fieldScm2 =
                        " Name = '"
                        + value.FieldIname
                        + "' StaticName='"
                        + value.FieldIname
                        + "' DisplayName='"
                        + value.FieldDname
                        + "' "
                        + value.FieldProp
                        + ">"
                        + "</Field>";
                }

                fieldScm += fieldScm2;

                let spField: SP.Field = spList.get_fields().addFieldAsXml(fieldScm, true, SP.AddFieldOptions.addFieldInternalNameHint);

                // インデックスを付ける列か
                if (value.IsIndex === true)
                {
                    colIndexFields.push(spField);
                }
            });

            // リストに列を追加する
            spList.update();
            await Common.SPExecuteQueryAsync(spClientContext, spList);

            await Common.SPExecuteArrayQueryAsync(spClientContext, colIndexFields);
            colIndexFields.forEach((value: SP.Field, index: number, array: SP.Field[]): void =>
            {
                value.set_indexed(true);
                value.update();
            });
            await Common.SPExecuteQueryAsync(spClientContext);

            return true;
        }
        catch (e)
        {
            return false;
        }
    }

    //
    // ファイル名を取得する
    //
    public static get_FileName(file: File): string
    {
        let aryFilePath: string[] = (<string>file.name).split(/\\/);
        if (aryFilePath.length > 0)
        {
            let fileName: string = aryFilePath[aryFilePath.length - 1];
            return (fileName === null || typeof fileName === "undefined") ? "" : fileName;
        }
        else
        {
            return "";
        }
    }

    //
    // テキスト形式でファイルを読み込む
    //
    public static ReadAsText(file: File): Promise<string>
    {
        return new Promise<string>((resolve: (value?: string) => void, reject: (reason?: SP.ClientRequestFailedEventArgs) => void): void =>
        {
            // FileReaderの作成
            let objFileReader: FileReader = new FileReader();

            // テキスト形式で読み込む
            objFileReader.readAsText(file, 'shift_jis');

            // 読み込み完了
            objFileReader.onload = (ev: Event): void =>
            {
                let strFileText = objFileReader.result;
                resolve(strFileText);
            };
            objFileReader.onerror = (ev: Event): void =>
            {
                reject();
            };
        });
    }

    //
    // バイナリ形式でファイルを読み込む
    //
    public static ReadAsBinary(file: File): Promise<any>
    {
        return new Promise<any>((resolve: (value?: string) => void, reject: (reason?: SP.ClientRequestFailedEventArgs) => void): void =>
        {
            // FileReaderの作成
            let objFileReader: FileReader = new FileReader();

            // バイナリ形式で読み込む
            objFileReader.readAsArrayBuffer(file);

            // 読み込み完了
            objFileReader.onload = (ev: Event): void =>
            {
                let strFileText: any = objFileReader.result;
                resolve(strFileText);
            };
            objFileReader.onerror = (ev: Event): void =>
            {
                reject();
            };
        });
    }

    //
    // ファイルをアップロードする
    //
    public static UploadFileToDocLib(arrayBuffer: any, serverRelativeUrlToFolder: string, fileName: string, serverUrl: string, siteUrl: string, appUrl: string): Promise<void>
    {
        return new Promise<void>((resolve: (value?: void) => void, reject: (reason?: SP.ClientRequestFailedEventArgs) => void): void =>
        {
            let executor: SP.RequestExecutor = new SP.RequestExecutor(appUrl);
            let content: string = Common.arrayBufferToBase64(arrayBuffer);
                
            // RESTのエンドポイント
            let fileCollectionEndpoint: string =
                appUrl
                + "/_api/sp.appcontextsite(@target)/web/getfolderbyserverrelativeurl('"
                + serverRelativeUrlToFolder
                + "')/files"
                + "/add(overwrite=true, url='"
                + fileName
                + "')?@target='"
                + siteUrl
                + "'";

            executor.executeAsync(
                {
                    url: fileCollectionEndpoint,
                    method: "POST",
                    binaryStringRequestBody: true,
                    body: content,
                    headers: {
                        "accept": "application/json;odata=verbose",
                        "content-length": arrayBuffer.byteLength
                    },
                    success: (data: any): void =>
                    {
                        resolve();
                    },
                    error: (data: any): void =>
                    {
                        resolve();
                    }
                });
        });
    }

    //
    // 与えられたキーワードクエリーで検索サービスにより検索した結果を返す
    //
    public static async Get_SPSearchExecutorResult(spClientContext: SP.ClientContext, keywordquery: string): Promise<SP.JsonObjectResult>
    {
        let objKeywordQuery: Microsoft.SharePoint.Client.Search.Query.KeywordQuery
            = new Microsoft.SharePoint.Client.Search.Query.KeywordQuery(spClientContext);
        objKeywordQuery.set_queryText(keywordquery);

        let objSearchExecutor: Microsoft.SharePoint.Client.Search.Query.SearchExecutor
            = new Microsoft.SharePoint.Client.Search.Query.SearchExecutor(spClientContext);

        // クエリーの作成
        let objResult: SP.JsonObjectResult = objSearchExecutor.executeQuery(objKeywordQuery);
        try
        {
            await Common.SPExecuteQueryAsync(spClientContext);
            // クエリーの結果を取得
            return objResult;
        }
        catch (e)
        {
            return null;
        }
    }


    //
    // メールを送信する
    //
    //  toUserList: メール送信先ユーザー
    //  subject: メール件名
    //  body: メール本文
    //
    public static SendEmail(spClientContext: SP.ClientContext, toUserList: SP.User[], subject: string, body: string, appUrl: string): Promise<void>
    {
        return new Promise<void>((resolve: (value?: void) => void, reject: (reason?: SP.ClientRequestFailedEventArgs) => void): void =>
        {
            let executor: SP.RequestExecutor = new SP.RequestExecutor(appUrl);

            // RESTのエンドポイント
            let restUrl: string = appUrl + "/_api/SP.Utilities.Utility.SendEmail";

            let toList: string[] = new Array<string>();
            toUserList.forEach((value: SP.User, index: number, array: SP.User[]): void =>
            {
                toList.push(value.get_loginName());
            });

            let mailObject =
            {
                "properties":
                {
                    "__metadata":
                    {
                        "type": "SP.Utilities.EmailProperties"
                    },
                    "To":
                    {
                        "results": toList
                    },
                    "Subject": subject,
                    "Body": body,
                    "AdditionalHeaders":
                    {
                        "__metadata":
                        {
                            "type": "Collection(SP.KeyValue)"
                        },
                        "results":
                        [
                            {
                                "__metadata": {
                                "type": 'SP.KeyValue'
                            },
                            "Key": "content-type",
                            "Value": 'text/html',
                            "ValueType": "Edm.String"
                            }
                        ]
                    }
                }   
            };

            let strMailObject: string = JSON.stringify(mailObject);
            let strBody: string = Common.sendMailReplacer(strMailObject);

            executor.executeAsync(
                {
                    url: restUrl,
                    method: "POST",
                    body: strBody,
                    headers: {
                        "Accept": "application/json;odata=verbose",
                        "Content-Type": "application/json;odata=verbose"
                    },
                    success: (data: any): void =>
                    {
                        resolve();
                    },
                    error: (data: any): void =>
                    {
                        resolve();
                    }
                });
        });
    }

    static sendMailReplacer(value: string)
    {
        let strValue: string = value.replace(/\\n/g, "<br />");
        return strValue;
    }


    //
    // 与えられた文字列の、\nを<br />に変換する
    //
    static CrToBrReplacer(value: string)
    {
        let strValue: string = value.replace(/\n/g, "<br />");
        return strValue;
    }

    //
    // バイナリデータをUTF8文字列にする
    //
    public static arrayBufferToBase64(buffer: any): string
    {
        let binary: string = ""
        let bytes: Uint8Array = new Uint8Array(buffer)
        let len: number = bytes.byteLength;
        for (let i: number = 0; i < len; i++)
        {
            binary += String.fromCharCode(bytes[i])
        }
        return binary;
    }

    //
    // ログインユーザーが、Webサイトに与えられた権限を有しているかを調べる
    //
    public static async DoesUserHaveWebPermissions(spClientContext: SP.ClientContext, spWeb: SP.Web, permission: SP.PermissionKind): Promise<boolean>
    {
        //let spBasePermisson: SP.BasePermissions = new SP.BasePermissions();
        //spBasePermisson.set(permission);
        //let usrPermission: SP.BooleanResult = spWeb.doesUserHavePermissions(spBasePermisson);

        //await Common.SPExecuteQueryAsync(spClientContext);
        //return usrPermission.get_value();
        await Common.SPExecutePropQueryAsync(spClientContext, spWeb, "EffectiveBasePermissions");
        let baseParmission: SP.BasePermissions = spWeb.get_effectiveBasePermissions();
        let hasPermission: boolean = baseParmission.has(permission);
        return hasPermission;
    }

    //
    // ログインユーザーが、オブジェクトに対して与えられた権限を有しているかを調べる
    //
    public static async DoesUserHaveListPermissions(spClientContext: SP.ClientContext, spList: SP.List, permission: SP.PermissionKind): Promise<boolean>
    {
        await Common.SPExecutePropQueryAsync(spClientContext, spList, "EffectiveBasePermissions");
        let baseParmission: SP.BasePermissions = spList.get_effectiveBasePermissions();
        await Common.SPExecuteQueryAsync(spClientContext);
        let hasPermission: boolean = baseParmission.has(permission);
        return hasPermission;
    }

    //
    // 与えられたユーザーが、オブジェクトに対して与えられた権限を有しているかを調べる
    //
    public static async DoesSpecifyUserHaveListPermissions(spClientContext: SP.ClientContext, spList: SP.List, spUser: SP.User, permission: SP.PermissionKind): Promise<boolean>
    {
        await Common.SPExecutePropQueryAsync(spClientContext, spList, "EffectiveBasePermissions");
        let loginName: string = spUser.get_loginName();
        let baseParmission: SP.BasePermissions = spList.getUserEffectivePermissions(loginName);
        await Common.SPExecuteQueryAsync(spClientContext);
        let hasPermission: boolean = baseParmission.has(permission);
        return hasPermission;
    }


    //
    // CSOMのexecuteQueryAsyncのawait用の実装
    //
    public static SPExecuteQueryAsync(spClientContext: SP.ClientContext, ...clientObjects: SP.ClientObject[]): Promise<void>
    {
        return new Promise<void>((resolve: (value?: void) => void, reject: (reason?: SP.ClientRequestFailedEventArgs) => void): void =>
        {
            clientObjects.forEach((value: SP.ClientObject, index: number, array: SP.ClientObject[]) =>
            {
                spClientContext.load(value);
            });

            spClientContext.executeQueryAsync(
                (sender: any, args: SP.ClientRequestSucceededEventArgs) =>
                {
                    resolve();
                },
                (sender: any, args: SP.ClientRequestFailedEventArgs) =>
                {
                    reject(args);
                });
        });
    }
    public static SPExecuteArrayQueryAsync(spClientContext: SP.ClientContext, clientObjects: SP.ClientObject[]): Promise<void>
    {
        return new Promise<void>((resolve: (value?: void) => void, reject: (reason?: SP.ClientRequestFailedEventArgs) => void): void =>
        {
            clientObjects.forEach((value: SP.ClientObject, index: number, array: SP.ClientObject[]) =>
            {
                spClientContext.load(value);
            });

            spClientContext.executeQueryAsync(
                (sender: any, args: SP.ClientRequestSucceededEventArgs) =>
                {
                    resolve();
                },
                (sender: any, args: SP.ClientRequestFailedEventArgs) =>
                {
                    reject(args);
                });
        });
    }
    public static SPExecutePropQueryAsync(spClientContext: SP.ClientContext, clientObject: SP.ClientObject, property: string): Promise<void>
    {
        return new Promise<void>((resolve: (value?: void) => void, reject: (reason?: SP.ClientRequestFailedEventArgs) => void): void =>
        {
            spClientContext.load(clientObject, property);

            spClientContext.executeQueryAsync(
                (sender: any, args: SP.ClientRequestSucceededEventArgs) =>
                {
                    resolve();
                },
                (sender: any, args: SP.ClientRequestFailedEventArgs) =>
                {
                    reject(args);
                });
        });
    }
    public static SPExecuteArryPropQueryAsync(spClientContext: SP.ClientContext, propArray: { clientObject: SP.ClientObject; property: string }[]): Promise<void>
    {
        return new Promise<void>((resolve: (value?: void) => void, reject: (reason?: SP.ClientRequestFailedEventArgs) => void): void =>
        {
            propArray.forEach((value: { clientObject: SP.ClientObject, property: string }, index: number, array: { clientObject: SP.ClientObject, property: string }[]) =>
            {
                spClientContext.load(value.clientObject, value.property);
            });

            spClientContext.executeQueryAsync(
                (sender: any, args: SP.ClientRequestSucceededEventArgs) =>
                {
                    resolve();
                },
                (sender: any, args: SP.ClientRequestFailedEventArgs) =>
                {
                    reject(args);
                });
        });
    }

    //
    // 指定msec時間待つ
    //
    public static Sleep(time: number): Promise<void>
    {
        return new Promise<void>((resolve: (value?: void) => void, reject: (reason?: SP.ClientRequestFailedEventArgs) => void): void =>
        {
            setTimeout((): void =>
            {
                resolve();
            },
            time);
        });
    }

    //
    // ブラウザの種類を取得する
    //
    public static GetBrowserKind(): BrowserKind
    {
        let retVal: BrowserKind = BrowserKind.Other;

        let userAgent: string = window.navigator.userAgent.toLowerCase();

        if ((userAgent.indexOf('msie') >= 0) && (userAgent.indexOf('trident/7.0') < 0))
        {
            retVal = BrowserKind.IE10Below;
        }
        else if (userAgent.indexOf('trident/7.0') >= 0)
        {
            retVal = BrowserKind.IE11;
        }
        else if (userAgent.indexOf('edge') >= 0)
        {
            retVal = BrowserKind.Edge;
        }
        else if (userAgent.indexOf('chrome') >= 0)
        {
            retVal = BrowserKind.Chrome;
        }
        else if (userAgent.indexOf('safari') >= 0)
        {
            retVal = BrowserKind.Safari;
        }
        else if (userAgent.indexOf('firefox') >= 0)
        {
            retVal = BrowserKind.FireFox;
        }
        else if (userAgent.indexOf('opera') >= 0)
        {
            retVal = BrowserKind.Opera;
        }
        else
        {
            retVal = BrowserKind.Other;
        }

        return retVal;
    }

    //
    // 画面を別Windowで開く
    //
    public static OpenWindow(strPage: string, strParam: string = "", strName: string): void
    {
        let spHostUrl: string = Common.get_ParameterByName("SPHostUrl");
        let url: string = strPage + "?SPHostUrl=" + spHostUrl + strParam;

        // モニター解像度
        let screenWidth: number = screen.width;
        let screenHeight: number = screen.height;
        let openWidth: number = Math.floor(screenWidth / 2);
        let openHeight: number = Math.floor(screenHeight / 2);
        let strOpenProp: string = "width=" + openWidth.toString() + ", height=" + openHeight.toString() + ", resizable=yes, menubar=no, toolbar=no, location=no, scrollbars=yes";

        window.open(url, strName, strOpenProp);
    }

    //
    // 画面を別Windowで開く
    //
    public static OpenWindowWithOutSPHostUrl(strPage: string, strParam: string = "", strName: string): void
    {
        let url: string = strPage + strParam;

        // モニター解像度
        let screenWidth: number = screen.width;
        let screenHeight: number = screen.height;
        let openWidth: number = Math.floor(screenWidth / 2);
        let openHeight: number = Math.floor(screenHeight / 2);
        let strOpenProp: string = "width=" + openWidth.toString() + ", height=" + openHeight.toString() + ", resizable=yes, menubar=no, toolbar=no, location=no, scrollbars=yes";

        window.open(url, strName, strOpenProp);
    }

    //
    // 画面を別タブで開く
    //
    public static OpenTabWindow(url: string): void
    {
        window.open(url, '_blank');
    }

    //
    // ドキュメントを開く
    //
    public static OpenDocument(url: string): void
    {
        window.location.href = url;
    }

    //
    // URLストリングを分解し、与えられたパラメータの値を返す
    //
    public static get_ParameterByName(name: string)
    {
        name = name.replace(/[\[]/, "\\\[").replace(/[\]]/, "\\\]");
        let regexS = "[\\?&]" + name + "=([^&#]*)";
        let regex = new RegExp(regexS);
        let results = regex.exec(window.location.href);
        if (results === null)
        {
            return "";
        }
        else
        {
            return decodeURIComponent(results[1].replace(/\+/g, " "));
        }
    }

    //
    // URLストリングを分解し、クエリーパラメータ文字列を返す
    //
    public static get_ParameterString()
    {
        let regexS = "\\?.*";
        let regex = new RegExp(regexS);
        let results = regex.exec(window.location.href);
        if (results === null)
        {
            return "";
        }
        else
        {
            return results[0];
        }
    }

    //
    // 言語に合わせて、日付をフォーマットする
    //
    public static get_MonthString(date: Date, language: string = "en-US"): string
    {
        if (date === null || typeof date === "undefined")
        {
            return "";
        }
        let fdate: Date = new Date(date);
        if (fdate === null || typeof fdate === "undefined")
        {
            return "";
        }
        if (language === "ja-JP")
        {
            return Common.get_FormatDate(fdate, "YYYY/MM");
        }
        else
        {
            return Common.get_FormatDate(fdate, "MM/YYYY");
        }
    }
    public static get_DateString(date: Date, language: string = "en-US"): string
    {
        if (date === null || typeof date === "undefined")
        {
            return "";
        }
        let fdate: Date = new Date(date);
        if (fdate === null || typeof fdate === "undefined")
        {
            return "";
        }
        if (language === "ja-JP")
        {
            return Common.get_FormatDate(fdate, "YYYY/MM/DD");
        }
        else
        {
            return Common.get_FormatDate(fdate, "MM/DD/YYYY");
        }
    }
    public static get_TimeString(date: Date): string
    {
        if (date === null || typeof date === "undefined")
        {
            return "";
        }
        let fdate: Date = new Date(date);
        if (fdate === null || typeof fdate === "undefined")
        {
            return "";
        }
        return Common.get_FormatDate(fdate, "hh:mm:ss");
    }
    public static get_FormatDate(date: Date, format: string = null): string
    {
        if (format === null || format === "" || typeof format === "undefined")
        {
            format = 'YYYY/MM/DD hh:mm:ss.SSS';
        }
        format = format.replace(/YYYY/g, date.getFullYear().toString());
        format = format.replace(/MM/g, ('0' + (date.getMonth() + 1).toString()).slice(-2));
        format = format.replace(/DD/g, ('0' + date.getDate().toString()).slice(-2));
        format = format.replace(/hh/g, ('0' + date.getHours().toString()).slice(-2));
        format = format.replace(/mm/g, ('0' + date.getMinutes().toString()).slice(-2));
        format = format.replace(/ss/g, ('0' + date.getSeconds().toString()).slice(-2));
        if (format.match(/S/g))
        {
            let milliSeconds = ('00' + date.getMilliseconds().toString()).slice(-3);
            let length = format.match(/S/g).length;
            for (let i = 0; i < length; i++)
            {
                format = format.replace(/S/, milliSeconds.substring(i, i + 1));
            }
        }
        return format;
    }

    //
    // GUIDを生成する
    //
    public static generateGUID(): string
    {
        let d = new Date().getTime();
        if (typeof performance !== 'undefined' && typeof performance.now === 'function')
        {
            d += performance.now(); //use high-precision timer if available
        }
        let sGUID: string = 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, (c: string): string =>
        {
            let r = (d + Math.random() * 16) % 16 | 0;
            d = Math.floor(d / 16);
            return (c === 'x' ? r : (r & 0x3 | 0x8)).toString(16);
        });

        return sGUID;
    }

    //
    // ファイル名から、OWAで表示するかを判定する
    //
    static DoseUseOWA(fileName: string): boolean 
    {
        let imageCss: string = "";
        let strExt: string = "";
        let useOwa: boolean = false;

        let aryFileNameSplit: string[] = fileName.split(".");
        let intAryLength: number = aryFileNameSplit.length;
        if (intAryLength > 1)
        {
            strExt = aryFileNameSplit[intAryLength - 1];
            strExt = strExt.toLowerCase();

            switch (strExt) 
            {
                case 'pptx':
                case 'ppt':
                case 'pptm':
                case 'docx':
                case 'doc':
                case 'docm':
                case 'xlsx':
                case 'xls':
                case 'xlsm':
                case 'one':
                    useOwa = true;
                    break;

                case 'pdf':
                case 'xsn':
                case 'js':
                case 'css':
                case 'stp':
                case 'zip':
                case 'txt':
                    useOwa = false;
                    break;

                default:
                    useOwa = false;
                    break;
            }
        }

        return useOwa;
    }

    //
    // サニタイジング
    //
    static HtmlEscape(str)
    {
        if (str === null || typeof str === "undefined")
        {
            return "";
        }
        return str.replace(/[<>&"'`]/g, (match) =>
        {
            const escape = {
                '<': '&lt;',
                '>': '&gt;',
                '&': '&amp;',
                '"': '&quot;',
                "'": '&#39;',
                '`': '&#x60;'
            };
            return escape[match];
        });
    }
}
