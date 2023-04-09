// ヒアリングフォルダーのID
const DRIVER_FOLDER_ID =
    PropertiesService.getScriptProperties().getProperty("DRIVER_FOLDER_ID");
// GoogleDrive 02.ヒアリング
const DRIVER = DriveApp.getFolderById(DRIVER_FOLDER_ID);
// ヒアリングアプリ 関連一覧のID
const HEARING_APP_LIST_ID =
    PropertiesService.getScriptProperties().getProperty("HEARING_APP_LIST_ID");
// アラートリスト
const ALERT_LIST_SHEET =
    SpreadsheetApp.openById(HEARING_APP_LIST_ID).getSheetByName("アラートリスト");
// ホワイトリスト
const WHITE_LIST_SHEET =
    SpreadsheetApp.openById(HEARING_APP_LIST_ID).getSheetByName("ホワイトリスト");
// ラストヒアリングリスト
const LAST_HEARING_LIST_SHEET =
    SpreadsheetApp.openById(HEARING_APP_LIST_ID).getSheetByName("ラストヒアリングリスト");
// 社員管理リスト スプレッドシートのシート
const MASTER_LIST_SHEET = SpreadsheetApp.openById(HEARING_APP_LIST_ID).getSheetByName("社員管理リスト");

// 社員管理リスト スプレッドシートのシート
const SELECT_ITEM_LIST = SpreadsheetApp.openById(HEARING_APP_LIST_ID).getSheetByName("選択アイテムリスト");
// 日時フォーマット
const DATE_FORMAT = "YYYY/MM/DD";

/**
 * メンバーの一覧を取得してhtmlのテーブル情報を作成する。
 * @returns json
 **/
function getMember(filters: Array<string> = ["", "", "", ""], alert: string = "", hearing: string = ""): Array<{}> {
    // 最終行
    const lastRow = MASTER_LIST_SHEET.getRange("A2:A")
        .getValues()
        .filter(String).length;
    // 一覧情報
    let list = MASTER_LIST_SHEET.getRange(2, 1, lastRow, 6).getValues();
    // アラート情報の最終行
    const lastAlertRow = ALERT_LIST_SHEET.getRange("A1:A")
        .getValues()
        .filter(String).length;
    // アラート情報
    const alertList = ALERT_LIST_SHEET.getRange(1, 1, lastAlertRow, 1).getValues();
    // 出力情報
    let result: Array<{}> = Array<{}>();

    // ヒアリングリスト最終行
    const hearingLastRow = LAST_HEARING_LIST_SHEET.getRange("A1:A")
        .getValues()
        .filter(String).length;
    // ヒアリング情報
    const hearingList = LAST_HEARING_LIST_SHEET.getRange(1, 1, hearingLastRow, 4).getValues()

    // 絞り込み
    for (let i = 0; i < filters.length; i++) {
        list = filterArray(list, i, filters[i])
    }
    // アラートの絞り込み
    list = filterAlert(list, alert);

    list = filterHearing(list, hearing);
    // 生成
    list.forEach(function (item) {
        const name: string = item[2] + item[0];
        let alert = alertList.filter(function (value) {
            return value[0] === name;
        });

        let doc = hearingList.filter(function (value) {
            return value[0] === name;
        });
        const user = {
            "name": name,
            "url": doc.length > 0 ? doc[0][2] : ""
        }
        const businessManager: string = item[1]
        let lastHearing = null
        if (doc.length > 0) {
            const now = setDayjs();
            const _date = setDayjs(doc[0][1])

            const isInterval: boolean = now.diff(_date, 'month') > 3
            lastHearing = {
                "date": _date.format(DATE_FORMAT),
                "url": doc[0][3],
                isInterval
            }
        }

        var department: string = item[3];
        if (department === "開発") {
            department += "(" + item[4] + ")";
        }
        else if (department.indexOf("(SI)") > -1) {
            const text = department.indexOf("開発") > -1 ? "(" + item[4] + ") : SI兼任" : " : SI兼任";
            department = department.replace("(SI)", text);
        }
        const isAlert: boolean = alert.length > 0
        result.push({
            user,
            businessManager,
            lastHearing,
            department,
            "alert": isAlert,
            "post": item[5]
        });
    });
    Logger.log(result);
    return result;
}

/**
 * ヒアリングの内容を保存する
 * @param name 対象者名
 * @param date ヒアリング実施日
 * @param alert アラート機能
 * @param comment ヒアリングメモ
 **/
function seve(
    name: string,
    date: string,
    alert: boolean,
    comment: string
): void {
    const _date = setDayjs(date);
    // 指定したディレクトリ名のディレクトリを取得
    const folderIterator = DRIVER.getFoldersByName(name);
    // 該当ディレクトリの情報
    let targetFolder: GoogleAppsScript.Drive.Folder;
    if (folderIterator.hasNext()) {
        // 存在する場合
        targetFolder = folderIterator.next();
    } else {
        targetFolder = DRIVER.createFolder(name);
    }
    let doc: GoogleAppsScript.Document.Document;

    const files = targetFolder.getFiles();
    const targetId: string = targetFolder.getUrl();
    while (files.hasNext()) {
        const file = files.next();
        if (file.getName() === _date.format(DATE_FORMAT)) {
            doc = DocumentApp.openById(file.getId());
            break;
        }
    }
    if (doc == undefined) {
        // ドキュメント生成
        doc = DocumentApp.create(_date.format(DATE_FORMAT));
    }

    const docId = doc.getId();
    doc.getBody().appendParagraph(comment);
    doc.saveAndClose();

    const docFile = DriveApp.getFileById(docId);
    targetFolder.addFile(docFile);

    changeAlert(name, alert);

    // ヒアリングリスト最終行
    const hearingLastRow = LAST_HEARING_LIST_SHEET.getRange("A1:A")
        .getValues()
        .filter(String).length;
    // ヒアリング情報
    const hearingList = LAST_HEARING_LIST_SHEET.getRange(1, 1, hearingLastRow, 4).getValues()
    let item = hearingList.filter(function (value) {
        return value[0] === name
    })

    if (item.length > 0) {
        hearingList.forEach(function (value) {
            if (value[0] == name) {
                value.splice(1, 1, _date.format(DATE_FORMAT))
                value.splice(3, 1, doc.getUrl())
            }
        });
    }
    else {
        hearingList.push([
            name,
            _date.format(DATE_FORMAT),
            targetFolder.getUrl(),
            doc.getUrl()
        ])
    }
    LAST_HEARING_LIST_SHEET.getRange(1, 1, hearingList.length, hearingList[0].length).setValues(hearingList);
}
/**
 * dayjsのセット関数
 * @param time 変換したい日時データ
 * @returns 日本時間で設定されたdayjs
 **/
function setDayjs(time: any = null): any {
    dayjs.dayjs.locale("ja");
    // 日時ライブラリ
    let date = dayjs.dayjs();
    if (time != null) {
        date = dayjs.dayjs(time);
    }
    return date;
}

/**
 * 自身のURLを取得する。
 * @returns 自身のURL
 */
function getAppUrl(): string {
    return ScriptApp.getService().getUrl();
}

/**
 * 最終ヒアリングメモのURLを取得
 * @param name 対象者
 * @returns 最終ヒアリングメモURL 未ヒアリングであれば空文字
 **/
function getLastDocumentMemo(name: string): any {
    // ヒアリングリスト最終行
    const hearingLastRow = LAST_HEARING_LIST_SHEET.getRange("A1:A")
        .getValues()
        .filter(String).length;
    // ヒアリング情報
    const hearingList = LAST_HEARING_LIST_SHEET.getRange(1, 1, hearingLastRow, 4).getValues()
    let doc = hearingList.filter(function (value) {
        return value[0] === name
    })
    return doc.length > 0 ? doc[0][3] : "";
}

/**
 * アラートの状態を変更する
 * @param name 対象
 * @param alert アラートのON/OFF
 **/
function changeAlert(name: string, alert: boolean): void {
    const lastAlertRow: number = ALERT_LIST_SHEET.getRange("A1:A")
        .getValues()
        .filter(String).length;
    // アラート情報
    let alertList = ALERT_LIST_SHEET.getRange(1, 1, lastAlertRow, 1).getValues();

    let result = alertList.filter(function (value) {
        return value[0] == name;
    });

    ALERT_LIST_SHEET.getRange("A1:A").clearContent();
    Logger.log(alert)
    // 配列に対象者名が無く、alertがtrueの場合
    if (result.length == 0 && alert) {
        alertList.push([name]);
    }

    // 配列にすでに名前が有りalertがfalseの場合
    if (result.length > 0 && !alert) {
        alertList.forEach(function (value) {
            if (value[0] == name) {
                Logger.log(name);
                Logger.log(value);
                alertList.splice(alertList.indexOf(value), 1);
            }
        })
    }
    ALERT_LIST_SHEET.getRange(
        1,
        1,
        alertList.length,
        alertList[0].length
    ).setValues(alertList);
}

/**
 * アラートの状態を取得
 * @param name 対象者名
 * @returns アラートの状態
 **/
function getAlert(name: string): boolean {
    const lastAlertRow: number = ALERT_LIST_SHEET.getRange("A1:A")
        .getValues()
        .filter(String).length;
    // アラート情報
    const alertList = ALERT_LIST_SHEET.getRange(1, 1, lastAlertRow, 1).getValues()

    let result = alertList.filter(function (value) {
        return value[0] === name;
    });
    return result.length > 0;
}

/**
 * チェックボックスの選択状態を切り替える
 * @param name 対象者名
 * @returns checked or 空文字
 **/
function checkedState(name) {
    return getAlert(name) ? "checked" : "";
}

/**
 * 閲覧可能なユーザーかチェック
 * @returns 閲覧可否
 **/
function checkWhiteUser(): boolean {
    //
    const effectiveUser = Session.getEffectiveUser().getEmail();
    const lastRow = WHITE_LIST_SHEET.getRange("A2:A")
        .getValues()
        .filter(String).length;
    const users = WHITE_LIST_SHEET.getRange(1, 1, lastRow, 1).getValues();
    let result = users.filter(function (value) {
        return value[0] === effectiveUser;
    });

    return result.length > 0;
}

/**
 * 配列にフィルタをかけて絞り込む
 * @param list 絞り込みたい配列
 * @param index 絞り込みたい内容の番号 
 * @param filter 絞り込み条件
 * @returns 絞り込んだ配列
 **/
function filterArray(list: Array<string[]>, index: number, filter: string): Array<string[]> {
    if (filter === "すべて" || filter === "") {
        return list
    }

    let result = list.filter(function (value) {
        return value[index].indexOf(filter) > -1
    })

    return result
}

/**
 * アラート状況で絞り込み
 * @param list 社員一覧
 * @param alert アラート状況
 * @returns 絞り込み結果
 */
function filterAlert(list: Array<string[]>, alert: string): Array<string[]> {
    let result: Array<string[]> = Array<string[]>();

    list.forEach(function (value) {
        const name = value[2] + value[0];
        if (alert === "あり" && getAlert(name)) {
            result.push(value);
        } else if (alert === "なし" && !getAlert(name)) {
            result.push(value);
        }
        else if (alert === "すべて") {
            result.push(value);
        }
    });
    return result;
}

/**
 * ヒアリング状況で絞り込み
 * @param list 社員一覧
 * @param hearing 絞り込み条件 
 * @returns 絞り込み結果
 */
function filterHearing(list: Array<string[]>, hearing: string): Array<string[]> {
    let result: Array<string[]> = Array<string[]>();
    // ヒアリングリスト最終行
    const hearingLastRow = LAST_HEARING_LIST_SHEET.getRange("AB:A")
        .getValues()
        .filter(String).length;
    // ヒアリング情報
    const hearingList = LAST_HEARING_LIST_SHEET.getRange(2, 1, hearingLastRow, 2).getValues();
    list.forEach(function (value) {
        const name = value[2] + value[0];
        const isInterval = getIsInterval(hearingList, name);
        const isHearing = getIsHearing(hearingList, name);
        Logger.log("");
        Logger.log("name : " + name);
        Logger.log("hearing : " + hearing);
        Logger.log("isInterval : " + isInterval);
        Logger.log("isHearing : " + isHearing );
        if (hearing === "すべて") {
            result.push(value);
        }
        else if (hearing === "未ヒアリング" && !isHearing) {
            result.push(value);
        } else if (isHearing) {
            if (hearing === "3ヶ月以上未実施" && isInterval) {
                result.push(value);
            }
            else if (hearing === "3ヶ月以内に実施" && !isInterval) {
                result.push(value);
            }
        }
    });
    return result;
}

/**
 * 最終ヒアリングから３ヶ月経っているかどうか
 * @param name 対象者名
 * @returns 結果
 */
function getIsInterval(hearingList: Array<string[]>, name: string): boolean {
    const now = setDayjs();
    let item = hearingList.filter(function (value) {
        return value[0] === name;
    })

    if (item.length === 0) {
        return false;
    }
    const _date = setDayjs(item[0][1]);
    Logger.log(now.diff(_date, 'month'));
    return now.diff(_date, 'month') > 3
}

function getIsHearing(hearingList: Array<string[]>, name: string): boolean {


    let item = hearingList.filter(function (value) {
        return value[0] === name;
    })

    return item.length > 0
}
/**
 * 会社リストを取得
 * @returns 会社リスト
 **/
function getCompany(): Array<string> {
    const companyLastRow = SELECT_ITEM_LIST.getRange("A2:A")
        .getValues()
        .filter(String).length;
    const companyList = Array.prototype.concat.apply([], SELECT_ITEM_LIST.getRange(2, 1, companyLastRow, 1).getValues());
    Logger.log(companyList);
    return companyList;
}

/**
 * 事業部リストを取得
 * @returns 事業部リスト
 */
function getDepartment(): Array<string> {
    const departmentLastRow = SELECT_ITEM_LIST.getRange("B2:B")
        .getValues()
        .filter(String).length;
    const departmentList = Array.prototype.concat.apply([], SELECT_ITEM_LIST.getRange(2, 2, departmentLastRow, 1).getValues());

    Logger.log(departmentList);
    return departmentList;
}

/**
 * 所属課リストを取得
 * @returns 所属課リスト
 +*/
function getDivision(): Array<string> {
    const divisionListLastRow = SELECT_ITEM_LIST.getRange("C2:C")
        .getValues()
        .filter(String).length;
    const divisionList = Array.prototype.concat.apply([], SELECT_ITEM_LIST.getRange(2, 3, divisionListLastRow, 1).getValues());


    Logger.log(divisionList);
    return divisionList;
}