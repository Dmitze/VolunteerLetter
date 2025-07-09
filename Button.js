function onOpen() {
  SpreadsheetApp.getActiveSpreadsheet().addMenu("📦 Меню", [
    { name: "🖨️ Створити волонтерські документи", functionName: "generateNextDocumentsBatch" }
  ]);
}
