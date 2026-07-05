const XlsxPopulate = require('xlsx-populate');
async function test() {
  const wb = await XlsxPopulate.fromBlankAsync();
  const sheet = wb.sheet(0);
  try {
    sheet.range(1, 1, 2, 2).merged(true);
    console.log("Range numbers work");
  } catch (e) {
    console.log("Range numbers fail:", e.message);
  }
}
test();
