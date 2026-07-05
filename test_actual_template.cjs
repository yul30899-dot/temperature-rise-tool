const XlsxPopulate = require('xlsx-populate');
async function test() {
  const wb = await XlsxPopulate.fromFileAsync('public/chart_template.xlsx');
  wb.sheet('原始数据').cell('A2').value('00:00:00');
  wb.sheet('原始数据').cell('B2').value(25.5);
  await wb.toFileAsync('public/test_export.xlsx');
  console.log("Done");
}
test();
