import xlsxwriter

# Create a workbook and add a thermalBatteryWorksheet.
workbook = xlsxwriter.Workbook('Material_Cost_Source.xlsx')
thermalBatteryWorksheet = workbook.add_worksheet()

# Add a bold format to use to highlight cells.
bold = workbook.add_format({'bold': True})

# Add a number format for cells with money.
money = workbook.add_format({'num_format': '$#,##0'})

# mm to foot conversion
mmFt = 304.8

# Write some data headers.
thermalBatteryWorksheet.write('A1', 'Material', bold)
thermalBatteryWorksheet.write('B1', 'Length (Feet)', bold)
thermalBatteryWorksheet.write('C1', 'Width (Feet)', bold)
thermalBatteryWorksheet.write('D1', 'Height (Feet)', bold)
thermalBatteryWorksheet.write('E1', 'Length (mm)', bold)
thermalBatteryWorksheet.write('F1', 'Width (mm)', bold)
thermalBatteryWorksheet.write('G1', 'Height (mm)', bold)
thermalBatteryWorksheet.write('H1', 'Cost', bold)
thermalBatteryWorksheet.write('I1', 'Quantity', bold)
thermalBatteryWorksheet.write('J1', 'Source', bold)

# Some data we want to write to the thermalBatteryWorksheet.
materials = (
	['EPS Foam', 2438.4, 1219.2, 152.4, 12.90, 68, 'google.com'],
)

# Start from the first cell below the headers.
row = 1
col = 0

# Iterate over the data and write it out row by row.
for material, lengthMm, widthMm, heightMm, cost, quantity, source in (materials):
    thermalBatteryWorksheet.write(row, col,     material)
    thermalBatteryWorksheet.write(row, col + 1, lengthMm / mmFt)
    thermalBatteryWorksheet.write(row, col + 2, widthMm / mmFt)
    thermalBatteryWorksheet.write(row, col + 3, heightMm / mmFt )
    thermalBatteryWorksheet.write(row, col + 4, lengthMm)
    thermalBatteryWorksheet.write(row, col + 5, widthMm)
    thermalBatteryWorksheet.write(row, col + 6, heightMm)
    thermalBatteryWorksheet.write(row, col + 7, cost)
    thermalBatteryWorksheet.write(row, col + 8, quantity)
    thermalBatteryWorksheet.write(row, col + 9, source)
    row += 1


# Write a total using a formula.
thermalBatteryWorksheet.write(row, 0, 'Total',       bold)
thermalBatteryWorksheet.write(row, 1, '=SUM(B2:B5)', money)


greenhouseWorksheet = workbook.add_worksheet()

# Write some data headers.
greenhouseWorksheet.write('A1', 'Material', bold)
greenhouseWorksheet.write('B1', 'Length (Feet)', bold)
greenhouseWorksheet.write('C1', 'Width (Feet)', bold)
greenhouseWorksheet.write('D1', 'Height (Feet)', bold)
greenhouseWorksheet.write('E1', 'Length (mm)', bold)
greenhouseWorksheet.write('F1', 'Width (mm)', bold)
greenhouseWorksheet.write('G1', 'Height (mm)', bold)
greenhouseWorksheet.write('H1', 'Cost', bold)
greenhouseWorksheet.write('I1', 'Quantity', bold)
greenhouseWorksheet.write('J1', 'Source', bold)

# Some data we want to write to the greenhouseWorksheet.
materials = (
	['EPS Foam', 2438.4, 609.6, 152.4, 1, 10, 'google.com'],
	['EPS Foam', 2438.4, 1219.2, 50.8, 1, 90, 'google.com'],
	['EPS Foam', 2438.4, 660.4, 25.4, 1, 30, 'google.com'],
)

# Start from the first cell below the headers.
row = 1
col = 0

# Iterate over the data and write it out row by row.
for material, lengthMm, widthMm, heightMm, cost, quantity, source in (materials):
    greenhouseWorksheet.write(row, col,     material)
    greenhouseWorksheet.write(row, col + 1, lengthMm / mmFt)
    greenhouseWorksheet.write(row, col + 2, widthMm / mmFt)
    greenhouseWorksheet.write(row, col + 3, heightMm / mmFt )
    greenhouseWorksheet.write(row, col + 4, lengthMm)
    greenhouseWorksheet.write(row, col + 5, widthMm)
    greenhouseWorksheet.write(row, col + 6, heightMm)
    greenhouseWorksheet.write(row, col + 7, cost)
    greenhouseWorksheet.write(row, col + 8, quantity)
    greenhouseWorksheet.write(row, col + 9, source)
    row += 1



workbook.close()
