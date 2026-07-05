const XlsxPopulate = require('xlsx-populate');
async function test() {
  const wb = await XlsxPopulate.fromBlankAsync();
  const sheet = wb.sheet(0);
  sheet.range("A1:C1").merged(true);
  sheet.cell("A1").value("Test");
  sheet.range("A1:C1").style("border", true);
  sheet.range("A1:C1").style("fill", "F2F2F2");
  sheet.range("A1:C1").style("horizontalAlignment", "center");
  sheet.range("A1:C1").style("verticalAlignment", "center");
  sheet.range("A1:C1").style("bold", true);
  sheet.range("A1:C1").style("fontFamily", "Microsoft YaHei");
  sheet.range("A1:C1").style("fontSize", 16);
  sheet.column("A").width(12);
  await wb.toFileAsync("style_test.xlsx");
  console.log("OK");
}
test();
