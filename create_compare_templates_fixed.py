import zipfile
import re
import os

def create_dual_chart_template(size):
    orig_file = f'public/chart_template_{size}.xlsx'
    final_file = f'public/chart_template_compare_{size}.xlsx'
    
    print(f"Creating {final_file} using pure zipfile method...")
    
    with zipfile.ZipFile(orig_file, 'r') as zorig:
        with zipfile.ZipFile(final_file, 'w') as zout:
            for item in zorig.infolist():
                content = zorig.read(item.filename)
                
                if item.filename == 'xl/charts/chart1.xml':
                    # write chart1
                    xml1 = content.decode('utf-8')
                    xml1 = xml1.replace(u"<a:t>温升曲线图</a:t>", u"<a:t>温升曲线图 (当前测试数据)</a:t>")
                    zout.writestr(item, xml1.encode('utf-8'))
                    
                    # create chart2
                    xml2 = content.decode('utf-8')
                    xml2 = xml2.replace(u"<a:t>温升曲线图</a:t>", u"<a:t>温升曲线图 (历史对比数据)</a:t>")
                    xml2 = xml2.replace(u"原始数据", u"对比原始数据")
                    xml2 = xml2.replace('"50010001"', '"50020001"')
                    xml2 = xml2.replace('"50010002"', '"50020002"')
                    zout.writestr('xl/charts/chart2.xml', xml2.encode('utf-8'))
                    
                elif item.filename == 'xl/worksheets/sheet2.xml':
                    # duplicate sheet2 as sheet3 for compare data
                    zout.writestr(item, content)
                    zout.writestr('xl/worksheets/sheet3.xml', content)
                
                elif item.filename == 'xl/workbook.xml':
                    xml = content.decode('utf-8')
                    new_sheet = '<sheet name="对比原始数据" sheetId="3" r:id="rId6"/>'
                    xml = xml.replace('</sheets>', new_sheet + '</sheets>')
                    zout.writestr(item, xml.encode('utf-8'))
                    
                elif item.filename == 'xl/_rels/workbook.xml.rels':
                    xml = content.decode('utf-8')
                    new_rel = '<Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>'
                    xml = xml.replace('</Relationships>', new_rel + '</Relationships>')
                    zout.writestr(item, xml.encode('utf-8'))
                    
                elif item.filename == 'docProps/app.xml':
                    xml = content.decode('utf-8')
                    # change <vt:i4>2</vt:i4> to 3
                    xml = re.sub(r'<vt:i4>2</vt:i4>', '<vt:i4>3</vt:i4>', xml)
                    # change size="2" to size="3" in TitlesOfParts
                    xml = re.sub(r'<TitlesOfParts><vt:vector size="2"', '<TitlesOfParts><vt:vector size="3"', xml)
                    # add sheet name
                    xml = xml.replace('</vt:vector></TitlesOfParts>', '<vt:lpstr>对比原始数据</vt:lpstr></vt:vector></TitlesOfParts>')
                    zout.writestr(item, xml.encode('utf-8'))
                
                elif item.filename == 'xl/drawings/drawing1.xml':
                    # duplicate the anchor
                    xml = content.decode('utf-8')
                    # find the only twoCellAnchor
                    anchor_match = re.search(r'<xdr:twoCellAnchor>.*?</xdr:twoCellAnchor>', xml, re.DOTALL)
                    anchor1 = anchor_match.group(0)
                    
                    # create anchor2 for the top chart (history data, rId2)
                    anchor2 = anchor1 # already at row 1 to 31
                    # change rel id
                    anchor2 = re.sub(r'r:id="rId\d+"', 'r:id="rId2"', anchor2)
                    # change drawing id/name to avoid conflict
                    anchor2 = re.sub(r'id="\d+" name="[^"]+"', 'id="3" name="Chart 2"', anchor2)
                    
                    # change anchor1 for the bottom chart (current data, rId1)
                    # move from 1->31 to 33->63
                    anchor1_mod = anchor1.replace('<xdr:row>1</xdr:row>', '<xdr:row>33</xdr:row>')
                    anchor1_mod = anchor1_mod.replace('<xdr:row>31</xdr:row>', '<xdr:row>63</xdr:row>')
                    
                    new_xml = xml.replace(anchor1, anchor2 + anchor1_mod)
                    zout.writestr(item, new_xml.encode('utf-8'))
                    
                elif item.filename == 'xl/drawings/_rels/drawing1.xml.rels':
                    xml = content.decode('utf-8')
                    # add relation for chart2
                    new_rel = '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart2.xml"/>'
                    new_xml = xml.replace('</Relationships>', new_rel + '</Relationships>')
                    zout.writestr(item, new_xml.encode('utf-8'))
                    
                elif item.filename == '[Content_Types].xml':
                    xml = content.decode('utf-8')
                    new_override_chart2 = '<Override PartName="/xl/charts/chart2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'
                    new_override_sheet3 = '<Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                    new_xml = xml.replace('</Types>', new_override_chart2 + new_override_sheet3 + '</Types>')
                    zout.writestr(item, new_xml.encode('utf-8'))
                    
                else:
                    zout.writestr(item, content)

    print(f"Created {final_file}")

if __name__ == '__main__':
    sizes = [200, 500, 1000, 2000, 3000, 5000]
    for s in sizes:
        create_dual_chart_template(s)
