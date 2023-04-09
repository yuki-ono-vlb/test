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
            .setFaviconUrl("https://drive.google.com/uc?id=1oC__fjaDQgupvA5v_CFRDHnB8Zo0K23x&.png")
            .addMetaTag('viewport', 'width=device-width, initial-scale=1')
            .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    const htmlIndex = HtmlService.createTemplateFromFile(page);
    for (const key in e.parameter) {
        htmlIndex[key] = e.parameter[key];
    }

    return htmlIndex
        .evaluate()
        .setTitle(page)
        .setFaviconUrl("https://drive.google.com/uc?id=1oC__fjaDQgupvA5v_CFRDHnB8Zo0K23x&.png")
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}