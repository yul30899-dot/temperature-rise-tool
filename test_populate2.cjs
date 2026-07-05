const XlsxPopulate = require('xlsx-populate');

async function testPopulate() {
    try {
        const wb = await XlsxPopulate.fromFileAsync('public/test_modified.xlsx');
        const sheet = wb.sheet('原始数据');
        sheet.cell('B1').value('CH89 Modified');
        sheet.cell('B2').value(100);
        await wb.toFileAsync('public/test_populated.xlsx');
        console.log("Successfully opened and saved populated file!");
    } catch (e) {
        console.error("XlsxPopulate Error:", e);
    }
}
testPopulate();
