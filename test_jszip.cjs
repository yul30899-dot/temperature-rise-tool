const JSZip = require('jszip');
const fs = require('fs');

async function test() {
    const arrayBuffer = fs.readFileSync('public/chart_template_200.xlsx');
    const zip = await JSZip.loadAsync(arrayBuffer);
    const chartFile = zip.file('xl/charts/chart1.xml');
    
    let xml = await chartFile.async('string');
    
    // We want to replace all <c:ser>...</c:ser> with our custom ones.
    const maxRows = 200;
    const channels = [
        {id: 89, name: 'T1'},
        {id: 90, name: 'T2'},
        {id: 91, name: 'T3'},
    ];
    
    function getColName(colIndex) {
        let temp, letter = '';
        while (colIndex > 0) {
            temp = (colIndex - 1) % 26;
            letter = String.fromCharCode(temp + 65) + letter;
            colIndex = (colIndex - temp - 1) / 26;
        }
        return letter;
    }

    let serNodes = '';
    channels.forEach((ch, idx) => {
        // Col B is 2
        const colLetter = getColName(idx + 2);
        
        serNodes += `<c:ser><c:idx val="${idx}"/><c:order val="${idx}"/><c:tx><c:strRef><c:f>原始数据!$${colLetter}$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>CH${ch.id}</c:v></c:pt></c:strCache></c:strRef></c:tx><c:spPr><a:ln w="28575"/></c:spPr><c:marker><c:symbol val="none"/></c:marker><c:cat><c:numRef><c:f>原始数据!$A$2:$A$${maxRows}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${maxRows-1}"/></c:numCache></c:numRef></c:cat><c:val><c:numRef><c:f>原始数据!$${colLetter}$2:$${colLetter}$${maxRows}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${maxRows-1}"/></c:numCache></c:numRef></c:val><c:smooth val="1"/></c:ser>`;
    });

    xml = xml.replace(/<c:ser>[\s\S]*?<\/c:ser>/g, '');
    // Insert back
    xml = xml.replace(/(<c:marker val="1"\/>|<c:axId)/, serNodes + '$1');
    
    zip.file('xl/charts/chart1.xml', xml);
    
    const outBuffer = await zip.generateAsync({ type: 'nodebuffer' });
    fs.writeFileSync('public/test_modified.xlsx', outBuffer);
    console.log("Created test_modified.xlsx");
}

test().catch(console.error);
