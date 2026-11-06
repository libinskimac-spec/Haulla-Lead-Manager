import csv
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# Read leads from CSV
with open("leads.csv", "r", newline="", encoding="utf-8") as file:
    rows = list(csv.reader(file))

# Create Excel workbook
workbook = Workbook()
sheet = workbook.active
sheet.title = "Haulla Leads"

for row in rows:
    sheet.append(row)

# Format header
for cell in sheet[1]:
    cell.font = Font(bold=True, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor="245A80")
    cell.alignment = Alignment(horizontal="center")

# Adjust column widths
for column in sheet.columns:
    width = max(len(str(cell.value or "")) for cell in column) + 3
    sheet.column_dimensions[
        get_column_letter(column[0].column)
    ].width = min(width, 40)

sheet.freeze_panes = "A2"
sheet.auto_filter.ref = sheet.dimensions

workbook.save("Haulla_Leads.xlsx")

print("Excel file created successfully!")
