# coding=utf-8
import openpyxl
import copy
from openpyxl.chart import Reference
from openpyxl.drawing.spreadsheet_drawing import TwoCellAnchor, AnchorMarker

def process(size):
    wb = openpyxl.load_workbook(f'public/chart_template_{size}.xlsx')
    
    ws_compare = wb.create_sheet(title=u"对比原始数据")
    ws_compare.cell(row=1, column=1, value=u"时间")
    for i in range(1, 201):
        ws_compare.cell(row=1, column=i+1, value=f"Compare_CH{i}")
        
    ws_summary = wb['温升数据']
    
    chart1 = ws_summary._charts[0]
    chart2 = copy.deepcopy(chart1)
    
    ws_summary._charts = []
    
    ws_summary.add_chart(chart2)
    ws_summary.add_chart(chart1)
    
    chart1.title = u"温升曲线图 (当前测试数据)"
    chart2.title = u"温升曲线图 (历史对比数据)"
    
    # Original is anchored from O2 to AG32
    # Chart2 (History) on top
    chart2.anchor = TwoCellAnchor(
        _from=AnchorMarker(col=14, colOff=0, row=1, rowOff=0),
        to=AnchorMarker(col=32, colOff=457200, row=31, rowOff=0)
    )
    # Chart1 (Current) on bottom (spaced below chart2)
    chart1.anchor = TwoCellAnchor(
        _from=AnchorMarker(col=14, colOff=0, row=33, rowOff=0),
        to=AnchorMarker(col=32, colOff=457200, row=63, rowOff=0)
    )
    
    wb.save(f'public/chart_template_compare_{size}.xlsx')
    print(f"Created public/chart_template_compare_{size}.xlsx")

if __name__ == "__main__":
    for size in [200, 500, 1000, 2000, 3000, 5000]:
        process(size)
