/**
 * Sparks Pitch Decks — Drive uploader (Google Apps Script web app)
 *
 * Runs as the account that deploys it and saves decks into
 * Pitch Decks / <Client> / <deck>.pptx. The build pipeline POSTs here.
 *
 * Deploy once: script.google.com → New project → paste this file → set SHARED_KEY →
 * Deploy → New deployment → type "Web app" → Execute as: Me → Who has access: Anyone
 * → Deploy → authorize → copy the Web app URL into the DRIVE_UPLOAD_URL environment
 * variable (and SHARED_KEY into DRIVE_UPLOAD_KEY). See drive/README.md.
 */
var ROOT_FOLDER_ID = "1bcAP65vMoyceh0UHXUFprdSmiOtIfTXf"; // "Pitch Decks"
var SHARED_KEY = "REPLACE-WITH-A-LONG-RANDOM-STRING";

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    if (!body.key || body.key !== SHARED_KEY) return out({ ok: false, error: "unauthorized" });
    if (!body.client || !body.name || !body.base64) return out({ ok: false, error: "client, name and base64 are required" });

    var root = DriveApp.getFolderById(ROOT_FOLDER_ID);
    var it = root.getFoldersByName(body.client);
    var folder = it.hasNext() ? it.next() : root.createFolder(body.client);

    // Replace an existing file of the same name (old copy goes to trash).
    var existing = folder.getFilesByName(body.name);
    while (existing.hasNext()) existing.next().setTrashed(true);

    var bytes = Utilities.base64Decode(body.base64);
    var blob = Utilities.newBlob(bytes, body.mimeType || "application/octet-stream", body.name);
    var file = folder.createFile(blob);
    return out({ ok: true, id: file.getId(), url: file.getUrl(), folder: folder.getName(), folderUrl: folder.getUrl() });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  }
}

function doGet() {
  return out({ ok: true, service: "sparks-pitch-decks-uploader" });
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
