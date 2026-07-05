const XlsxPopulate = require('xlsx-populate');
async function test() {
  const wb = await XlsxPopulate.fromFileAsync('test_chart.xlsx');
  wb.sheet('RawData').cell('B3').value(999);
  await wb.toFileAsync('test_chart_modified_populate.xlsx');
  console.log("Done xlsx-populate");
}
test();
