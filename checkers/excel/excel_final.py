from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font
from openpyxl.chart import PieChart, Reference, BarChart

# Load the provided dataset
workbook = load_workbook('Final Project Dataset.xlsx')
sheet = workbook.active
sheet.title = "Data Analysis"

# Ensure columns are labeled correctly
column_headers = ["Employee ID", "Name", "Department", "Training Hours", "Tasks Completed", "Digital Skills", "Column G", "Column H", "Column I", "Column J"]
for col_num, header in enumerate(column_headers, start=1):
    sheet.cell(row=1, column=col_num).value = header

# Calculate averages for columns C to J
for col_num in range(3, 11):
    avg_formula = f"=AVERAGE({chr(64 + col_num)}2:{chr(64 + col_num)}18)"
    sheet.cell(row=19, column=col_num).value = avg_formula

# Calculate total training hours and tasks completed
sheet.cell(row=19, column=4).value = "=SUM(D2:D18)"  # Training Hours
sheet.cell(row=19, column=5).value = "=SUM(E2:E18)"  # Tasks Completed

# Apply IF statement in column K for Digital Skills
for row in range(2, 19):
    sheet.cell(row=row, column=11).value = f"=IF(F{row}>=6, \"No training is needed\", \"Need to take training\")"

# Create charts on Sheet 1
chart_sheet = workbook["Data Analysis"]

# Bar chart for Training Hours
bar_chart = BarChart()
data = Reference(chart_sheet, min_col=4, min_row=1, max_row=18)
categories = Reference(chart_sheet, min_col=2, min_row=2, max_row=18)
bar_chart.add_data(data, titles_from_data=True)
bar_chart.set_categories(categories)
bar_chart.title = "Training Hours by Employee"
chart_sheet.add_chart(bar_chart, "L1")

# Pie chart for Tasks Completed
pie_chart = PieChart()
data = Reference(chart_sheet, min_col=5, min_row=1, max_row=18)
categories = Reference(chart_sheet, min_col=2, min_row=2, max_row=18)
pie_chart.add_data(data, titles_from_data=True)
pie_chart.set_categories(categories)
pie_chart.title = "Tasks Completed Distribution"
chart_sheet.add_chart(pie_chart, "L16")

# Create a new sheet for Department Analysis
dept_sheet = workbook.create_sheet("Department Analysis")
dept_sheet.append(["Department", "Number of Employees"])

# Calculate number of employees by department
departments = {}
for row in range(2, 19):
    dept = sheet.cell(row=row, column=3).value
    departments[dept] = departments.get(dept, 0) + 1

for dept, count in departments.items():
    dept_sheet.append([dept, count])

# Create a pie chart for Department Analysis
pie_chart_dept = PieChart()
data = Reference(dept_sheet, min_col=2, min_row=2, max_row=len(departments) + 1)
categories = Reference(dept_sheet, min_col=1, min_row=2, max_row=len(departments) + 1)
pie_chart_dept.add_data(data, titles_from_data=True)
pie_chart_dept.set_categories(categories)
pie_chart_dept.title = "Employees by Department"
dept_sheet.add_chart(pie_chart_dept, "E5")

# Save the workbook
workbook.save('Final_Project_Excel.xlsx')
