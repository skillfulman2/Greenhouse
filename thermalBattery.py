import xlsxwriter

# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('Expenses02.xlsx')
worksheet = workbook.add_worksheet()

# Add a bold format to use to highlight cells.
bold = workbook.add_format({'bold': True})

# Add a number format for cells with money.
money = workbook.add_format({'num_format': '$#,##0'})

# Write some data headers.
worksheet.write('A1', 'Material', bold)
worksheet.write('B1', 'Length (Feet)', bold)
worksheet.write('C1', 'Width (Feet)', bold)
worksheet.write('D1', 'Height (Feet)', bold)
worksheet.write('E1', 'Length (mm)', bold)
worksheet.write('F1', 'Width (mm)', bold)
worksheet.write('G1', 'Height (mm)', bold)
worksheet.write('H1', 'Cost', bold)
worksheet.write('I1', 'Quantity', bold)
worksheet.write('J1', 'Source', bold)

# Some data we want to write to the worksheet.
expenses = (
	['EPS Foam', 4, 8, .5, ,20, 50, 'google.com'],
	['EPS Foam', 4, 8, .5, ,20, 50, 'google.com'],
	['EPS Foam', 4, 8, .5, ,20, 50, 'google.com'],
	['EPS Foam', 4, 8, .5, ,20, 50, 'google.com'],
	['EPS Foam', 4, 8, .5, ,20, 50, 'google.com'],
)

# Start from the first cell below the headers.
row = 1
col = 0

# Iterate over the data and write it out row by row.
for material, lengthFt, widthFt, heightFt, cost, quantity, source in (expenses):
    worksheet.write(row, col,     material)
    worksheet.write(row, col + 1, lengthFt)
    worksheet.write(row, col + 2, widthFt)
    worksheet.write(row, col + 3, heightFt)
    worksheet.write(row, col + 4, quantity)
    worksheet.write(row, col + 5, quantity)
    worksheet.write(row, col + 6, quantity)
    worksheet.write(row, col + 7, cost)
    worksheet.write(row, col + 8, cost)
    worksheet.write(row, col + 9, cost)
    row += 1


# Write a total using a formula.
worksheet.write(row, 0, 'Total',       bold)
worksheet.write(row, 1, '=SUM(B2:B5)', money)

workbook.close()
