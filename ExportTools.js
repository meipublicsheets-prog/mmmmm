function exportSpreadsheetToExcelAndEmail() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ssId = ss.getId();
    const ssName = ss.getName();
    const userEmail = Session.getActiveUser().getEmail();
    const url = "https://docs.google.com/spreadsheets/d/" + ssId + "/export?format=xlsx";
    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(url, {
        headers: {
            'Authorization': 'Bearer ' + token
        }
    });

    const excelBlob = response.getBlob().setName(ssName + ".xlsx");
    let folder;
    try {
        const folderId = (typeof FOLDERS !== 'undefined' && FOLDERS.IMS_Reports) ? FOLDERS.IMS_Reports : '';
        if (folderId) {
            folder = DriveApp.getFolderById(folderId);
        } else {
            // Fallback: create in root or find by name
            const folders = DriveApp.getFoldersByName("IMS Reports");
            if (folders.hasNext()) {
                folder = folders.next();
            } else {
                folder = DriveApp.createFolder("IMS Reports");
            }
        }

        const file = folder.createFile(excelBlob);
        Logger.log("Excel backup saved: " + file.getUrl());

    } catch (e) {
        Logger.log("Error saving Excel backup: " + e.toString());
    }
    try {
        const subject = "IMS System Export - " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
        const body = "Attached is the latest XLSX export of the " + ssName + " workspace.\n\n" +
            "Exported by: " + userEmail + "\n" +
            "Timestamp: " + new Date().toString();

        GmailApp.sendEmail(userEmail, subject, body, {
            attachments: [excelBlob]
        });

        SpreadsheetApp.getUi().alert("Export successful! A copy has been emailed to " + userEmail + " and saved to the Reports folder.");

    } catch (e) {
        Logger.log("Error sending email: " + e.toString());
        SpreadsheetApp.getUi().alert("Export created but email failed: " + e.toString());
    }
}
