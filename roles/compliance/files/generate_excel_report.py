#!/usr/bin/env python3
import json, os, glob
import pandas as pd

report_dir = "/tmp/compliance_report"
data_files = glob.glob(os.path.join(report_dir, "*.json"))

rows = []
for f in data_files:
    with open(f) as infile:
        rows.append(json.load(infile))

df = pd.DataFrame(rows)
df = df[["ip", "name", "os", "kernel", "uptime", "compliance"]]

excel_file = os.path.join(report_dir, "compliance_report.xlsx")
with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
    df.to_excel(writer, index=False, sheet_name='Compliance Report')

    # Formatting (eye-catching)
    workbook = writer.book
    worksheet = writer.sheets['Compliance Report']
    for col_cells in worksheet.columns:
        length = max(len(str(cell.value)) for cell in col_cells)
        worksheet.column_dimensions[col_cells[0].column_letter].width = length + 4
    header_font = openpyxl.styles.Font(bold=True, color="FFFFFF")
    fill = openpyxl.styles.PatternFill("solid", fgColor="4CAF50")
    for cell in worksheet[1]:
        cell.font = header_font
        cell.fill = fill
