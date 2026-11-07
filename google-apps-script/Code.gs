const CONFIG = {
  ROOT_FOLDER_ID: "1aZYisjIWGu_iclvVTCyHjxDDVndryG9P",
  SPREADSHEET_ID: "1kleyK8-TeE5mxIDg-KIxG9K3qNOSRWns1IPFda00g6U",
  SHEET_NAME: "PDI History"
};

function setup() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sh = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sh) sh = ss.insertSheet(CONFIG.SHEET_NAME);
  if (sh.getLastRow() === 0) sh.appendRow([
    "Timestamp","Record ID","Inspector","Inspection Date","Inspection Line",
    "Inspection Lot","Shift","Total Sample QTY","Total NG","Total QTY",
    "File Name","File URL","Folder URL","Items JSON"
  ]);
  DriveApp.getFolderById(CONFIG.ROOT_FOLDER_ID).getName();
  return "Setup OK";
}

function doGet(e) {
  return json_({success:true, message:"PDI Inspection Apps Script is running"});
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    if (body.action !== "save") throw new Error("Unsupported action");
    return json_(saveInspection_(body));
  } catch (err) {
    return json_({success:false,message:String(err && err.message || err)});
  }
}

function saveInspection_(body) {
  if (!body.fileName || !body.base64 || !body.inspection) throw new Error("Payload file tidak lengkap");
  const data = body.inspection;
  const date = String(data.inspectionDate || "");
  const m = date.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) throw new Error("Inspection Date harus YYYY-MM-DD");
  const root = DriveApp.getFolderById(CONFIG.ROOT_FOLDER_ID);
  const yearFolder = getOrCreateFolder_(root, m[1]);
  const monthFolder = getOrCreateFolder_(yearFolder, m[2]);
  const dayFolder = getOrCreateFolder_(monthFolder, m[3]);
  const bytes = Utilities.base64Decode(body.base64);
  const blob = Utilities.newBlob(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", body.fileName);
  const file = dayFolder.createFile(blob);
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sh = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sh) { setup(); sh = ss.getSheetByName(CONFIG.SHEET_NAME); }
  sh.appendRow([
    new Date(), data.id || "", data.inspector || "", date, (data.line || []).join(", "),
    data.lotDate || "", data.shift || "", data.totalSample || 0, data.totalNg || 0,
    data.totalQty || 0, body.fileName, file.getUrl(), dayFolder.getUrl(),
    JSON.stringify(data.items || [])
  ]);
  return {
    success:true, fileId:file.getId(), fileUrl:file.getUrl(),
    downloadUrl:"https://drive.google.com/uc?export=download&id="+file.getId(),
    folderUrl:dayFolder.getUrl(), fileName:body.fileName
  };
}

function getOrCreateFolder_(parent, name) {
  const found = parent.getFoldersByName(name);
  return found.hasNext() ? found.next() : parent.createFolder(name);
}
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}