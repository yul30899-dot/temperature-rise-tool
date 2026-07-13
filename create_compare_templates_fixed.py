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
                    zout.writestr(item, content)
                    # create chart2
                    xml2 = content.decode('utf-8')
                    xml2 = xml2.replace(u"<a:t>温升曲线图</a:t>", u"<a:t>温升曲线图 (历史对比数据)</a:t>")
                    xml2 = xml2.replace('"50010001"', '"50020001"')
                    xml2 = xml2.replace('"50010002"', '"50020002"')
                    zout.writestr('xl/charts/chart2.xml', xml2.encode('utf-8'))
                
                elif item.filename == 'xl/drawings/drawing1.xml':
                    # duplicate the anchor
                    xml = content.decode('utf-8')
                    # find the only twoCellAnchor
                    anchor_match = re.search(r'<xdr:twoCellAnchor>.*?</xdr:twoCellAnchor>', xml, re.DOTALL)
                    anchor1 = anchor_match.group(0)
                    
                    # create anchor2 for the top chart (history data)
                    anchor2 = anchor1.replace('<xdr:row>32</xdr:row>', '<xdr:row>1</xdr:row>') # from row 1 (0-indexed)
                    anchor2 = anchor2.replace('<xdr:row>62</xdr:row>', '<xdr:row>31</xdr:row>') # to row 31 (0-indexed)
                    # change rel id
                    anchor2 = re.sub(r'r:id="rId\d+"', 'r:id="rId2"', anchor2)
                    
                    # keep anchor1 for the bottom chart (current data)
                    # already positioned at 32 to 62
                    anchor1_mod = anchor1
                    
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
                    new_override = '<Override PartName="/xl/charts/chart2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'
                    new_xml = xml.replace('</Types>', new_override + '</Types>')
                    zout.writestr(item, new_xml.encode('utf-8'))
                    
                else:
                    zout.writestr(item, content)

    print(f"Created {final_file}")

if __name__ == '__main__':
    sizes = [200, 500, 1000, 2000, 3000, 5000]
    for s in sizes:
        create_dual_chart_template(s)
