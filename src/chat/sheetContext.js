/* global Excel */

const MAX_ROWS = 20;
const MAX_COLS = 8;

function trimValues(values) {
  const rows = Array.isArray(values) ? values.slice(0, MAX_ROWS) : [];
  return rows.map((row) => (Array.isArray(row) ? row.slice(0, MAX_COLS) : []));
}

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

    const values = trimValues(range.values);
    const address = range.address || "";
    return {
      available: true,
      sheet: sheet.name,
      address,
      rowCount: range.rowCount,
      columnCount: range.columnCount,
      values,
      summary: `${sheet.name}!${address} (${range.rowCount} x ${range.columnCount})`,
    };
  });
}
