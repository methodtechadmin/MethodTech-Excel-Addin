/* global Excel */

export async function readSheetContext() {
  if (typeof Excel === "undefined") {
    return { available: false, summary: "Open this pane inside Excel to read the sheet." };
  }

  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = context.workbook.getSelectedRange();
    sheet.load("name");
    range.load("address, values, rowCount, columnCount");
    await context.sync();

    const values = Array.isArray(range.values) ? range.values : [];
    const address = range.address || "";
    const location = address.includes("!") ? address : `${sheet.name}!${address}`;
    return {
      available: true,
      sheet: sheet.name,
      address,
      rowCount: range.rowCount,
      columnCount: range.columnCount,
      values,
      summary: `${location} (${range.rowCount} x ${range.columnCount})`,
    };
  });
}
