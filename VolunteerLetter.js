function generateNextDocumentsBatch() {
  const templateDocId = "";
  const folderId      = "";
  const sheetName     = "VolunteerLetter";
  const statusColumn  = 8;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const data  = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
  const statuses = sheet.getRange(2, statusColumn, data.length).getValues();

  let created = 0;
  const maxToCreate = 5;

  for (let i = 0; i < data.length && created < maxToCreate; i++) {
    const row = data[i];
    const status = statuses[i][0];

    if (status === "✅ Створено") continue;
    if (row.some(cell => cell === "" || cell === null)) continue;

    const timestamp = formatDate(row[0]);
    const fullName  = row[1];
    const birthDate = formatDate(row[2]);
    const ipn       = row[3];
    const passport  = row[4];
    const issueDate = formatDate(row[5]);
    const issuedBy  = row[6];

    const docName   = `${timestamp}_${fullName}`.replace(/\s+/g, "_") + ".docx";

    // Перевірка, чи файл уже є в папці
    if (fileExists(docName, folderId)) {
      sheet.getRange(i + 2, statusColumn).setValue("✅ Створено (вже існує)");
      continue;
    }

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

      sheet.getRange(i + 2, statusColumn).setValue("✅ Створено");
      created++;
    } catch (e) {
      sheet.getRange(i + 2, statusColumn).setValue("❌ Помилка");
      Logger.log("❌ Створення документа для рядка " + (i + 2) + ": " + e);
    }
  }

  const msg = created
    ? `✅ Створено ${created} документ(ів)`
    : "ℹ️ Немає нових рядків для обробки або документи вже існують";
  SpreadsheetApp.getUi().alert(msg);
}

function fileExists(fileName, folderId) {
  const folder = DriveApp.getFolderById(folderId);
  const files  = folder.getFilesByName(fileName);
  return files.hasNext();
}

function formatDate(value) {
  if (!value) return "__.__.__";
  let date;

  if (typeof value === "string") {
    const parts = value.split(/[./-]/);
    if (parts.length === 3) {
      const year = parseInt(parts[2]) < 100 ? 2000 + parseInt(parts[2]) : parseInt(parts[2]);
      date = new Date(year, parseInt(parts[1]) - 1, parseInt(parts[0]));
    } else {
      date = new Date(value);
    }
  } else if (typeof value === "number") {
    date = new Date((value - 25569) * 86400 * 1000);
  } else if (value instanceof Date) {
    date = value;
  }

  if (isNaN(date.getTime())) return "__.__.__";
  return `${date.getDate().toString().padStart(2, "0")}.${(date.getMonth()+1).toString().padStart(2, "0")}.${date.getFullYear()}`;
}


