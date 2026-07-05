const ExcelJS = require('exceljs');
async function test() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile('test_chart.xlsx');
  const ws = wb.getWorksheet('RawData');
  ws.getCell('B3').value = 999;
  await wb.xlsx.writeFile('test_chart_modified.xlsx');
  console.log("Done");
}
test();
