import xlsxwriter

def create_template(max_rows):
    filename = f"public/chart_template_{max_rows}.xlsx"
    workbook = xlsxwriter.Workbook(filename)
    ws_summary = workbook.add_worksheet("温升数据")
    ws_raw = workbook.add_worksheet("原始数据")

    # Populate header
    ws_raw.write_string(0, 0, "Time")
    for i in range(1, 91):
        ws_raw.write_string(0, i, f"CH{i}")

    # Create chart - classic Line chart
    chart = workbook.add_chart({'type': 'line'})

    # Add series directly (no named ranges, robust)
    for i in range(1, 91):
        chart.add_series({
            'name':       ['原始数据', 0, i],
            'categories': ['原始数据', 1, 0, max_rows - 1, 0],
            'values':     ['原始数据', 1, i, max_rows - 1, i],
            'line':       {'width': 2.25}, # standard Excel thick line
            'smooth':     True
        })

    # Configure axes
    chart.set_title({'name': '温升曲线图', 'name_font': {'size': 14, 'bold': True}})
    
    chart.set_x_axis({
        'name': '时间',
        'label_position': 'next_to',
        'num_font': {'rotation': -90}, # Rotate text vertically so Excel auto-packs them densely
        'interval_unit': 1, # FORCE Excel to attempt to draw every label we provide
    })
    
    chart.set_y_axis({
        'name': '温度 (℃)',
        'major_gridlines': {'visible': True},
        'label_position': 'next_to',
        'major_unit': 10, # Set temperature interval to 10 as requested
    })
    
    # Place legend at the bottom
    chart.set_legend({
        'position': 'bottom',
        'font': {'size': 11} # larger font size for channels
    })

    # Set chart dimensions (pixels)
    chart.set_size({'width': 1200, 'height': 600})

    # Insert chart
    ws_summary.insert_chart('O2', chart)

    workbook.close()
    print(f"Created template for {max_rows} rows.")

if __name__ == "__main__":
    for r in [200, 500, 1000, 2000, 3000, 5000]:
        create_template(r)
