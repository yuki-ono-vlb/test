// get
function doGet(e): GoogleAppsScript.Content.TextOutput | GoogleAppsScript.HTML.HtmlOutput {
    let page = e.parameter.page

    Logger.log(page)
    if(!page){
        Logger.log("index");
        page = "index";
    }

    Logger.log(page)

    if(page === "json"){
        const payload = JSON.stringify(getMember());
        ContentService.createTextOutput()
        const output = ContentService.createTextOutput();
        output.setMimeType(ContentService.MimeType.JSON);
        output.setContent(payload);
        // return response-data
        return output;
    }

    if (!checkWhiteUser()) {
        const blockHtml = HtmlService.createTemplateFromFile("block");
        blockHtml.result = "閲覧できません。";
        return blockHtml
            .evaluate()
            .setTitle("閲覧禁止")
            .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    const htmlIndex = HtmlService.createTemplateFromFile(page);
    for (const key in e.parameter) {
        htmlIndex[key] = e.parameter[key];
    }

    return htmlIndex
        .evaluate()
        .setTitle(page)
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
// post
function doPost(e): GoogleAppsScript.HTML.HtmlOutput {
    if (!checkWhiteUser()) {
        const blockHtml = HtmlService.createTemplateFromFile("block");
        blockHtml.result = "書き込みできません。";
        return blockHtml
            .evaluate()
            .setTitle("書き込み禁止")
            .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
    Logger.log(e);

    const htmlIndex = HtmlService.createTemplateFromFile(e.parameter.page);
    const name = e.parameter.name;
    htmlIndex["name"] = name;

    const comment = e.parameter.comment;
    const date = e.parameter.date;
    const alert = e.parameter.alert === null ? false : e.parameter.alert;
    const page = e.parameter.page
    if (page === "memo") {
        seve(name, date, alert, comment);
    }
}
