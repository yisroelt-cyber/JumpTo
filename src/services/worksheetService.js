// src/services/worksheetService.js

/* global Excel */

export async function getVisibleWorksheets() {
  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/id,name,visibility");

    await context.sync();

    return sheets.items
      .filter(ws => ws.visibility === Excel.SheetVisibility.visible)
      .map(ws => ({
        id: ws.id,
        name: ws.name
      }));
  });
}

export async function activateWorksheetById(worksheetId) {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(worksheetId);
    sheet.activate();
    await context.sync();
  });
}
