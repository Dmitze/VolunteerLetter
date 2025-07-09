function onFormSubmit(e) {
  const templateDocId = "";
  const folderId      = "";
  const sheetName     = "VolunteerLetter";
  const statusColumn  = 8;
    if (!e || !e.values) {
    Logger.log("⚠️ Подія не має даних e.values");
    return;
  }
  const row = e.values.slice(0, 7); // A-G
  if (row.some(cell => cell === "" || cell === null)) return;

  const timestamp = formatDate(row[0]);
  const fullName  = row[1];
  const birthDate = formatDate(row[2]);
  const ipn       = row[3];
  const passport  = row[4];
  const issueDate = formatDate(row[5]);
  const issuedBy  = row[6];

  const docName = `${timestamp}_${fullName}`.replace(/\s+/g, "_") + ".docx";

  if (fileExists(docName, folderId)) return;

  try {
    const docFile = DriveApp.getFileById(templateDocId).makeCopy(docName);
    DriveApp.getFolderById(folderId).addFile(docFile);

    const doc = DocumentApp.openById(docFile.getId());

    doc.replaceText("\\{ПІБ\\}", fullName);
    doc.replaceText("\\{Дата народження\\}", birthDate);
    doc.replaceText("\\{ІПН\\}", ipn);
    doc.replaceText("\\{Серія та номер паспорта\\}", passport);
    doc.replaceText("\\{Дата видачі паспорта\\}", issueDate);
    doc.replaceText("\\{Ким видано\\}", issuedBy);
    doc.saveAndClose();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, statusColumn).setValue("✅ Створено");
  } catch (err) {
    Logger.log("❌ FormSubmit Error: " + err);
  }
}
