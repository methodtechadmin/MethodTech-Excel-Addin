/* global Excel */

function chartType(name) {
  if (typeof Excel === "undefined") {
    return null;
  }
  if (String(name || "line").toLowerCase() === "line") {
    return Excel.ChartType.line;
  }
  return null;
}

async function createChart(context, action) {
  const type = chartType(action.chart_type);
  if (!type) {
    throw new Error("Only a line chart is supported in this step.");
  }
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = action.data_range
    ? sheet.getRange(action.data_range)
    : context.workbook.getSelectedRange();
  const chart = sheet.charts.add(type, range, Excel.ChartSeriesBy.auto);
  if (action.title) {
    chart.title.text = String(action.title);
  }
  chart.setPosition("A12", "H28");
}

async function writeRange(context, action) {
  const start = String(action.start_cell || "A1");
  const values = Array.isArray(action.values) ? action.values : [];
  if (!values.length || !Array.isArray(values[0])) {
    throw new Error("write_range needs rows of values.");
  }
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const target = sheet.getRange(start).getResizedRange(values.length - 1, values[0].length - 1);
  target.values = values;
}

export async function applyExcelActions(actions) {
  const list = Array.isArray(actions) ? actions : [];
  if (!list.length) {
    return [];
  }
  if (typeof Excel === "undefined") {
    throw new Error("Excel actions run inside the workbook.");
  }

  const done = [];
  await Excel.run(async (context) => {
    for (const action of list) {
      const name = action && action.action;
      if (name === "create_chart") {
        await createChart(context, action);
        done.push("line chart");
      } else if (name === "write_range") {
        await writeRange(context, action);
        done.push(`values at ${action.start_cell || "A1"}`);
      }
    }
    await context.sync();
  });
  return done;
}
