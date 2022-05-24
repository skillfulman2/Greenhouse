import xlsxwriter

# Create a workbook and add a thermalBatteryWorksheet.
workbook = xlsxwriter.Workbook('../excelWorkbook/Material_Cost_Source.xlsx')
thermalBatteryWorksheet = workbook.add_worksheet()

# Add a bold format to use to highlight cells.
bold = workbook.add_format({'bold': True})

# Add a number format for cells with money.
money = workbook.add_format({'num_format': '$#,##0'})

# Inch to mm conversion
i2m = 25.4


# Write some data headers.
thermalBatteryWorksheet.write('A1', 'Material', bold)
thermalBatteryWorksheet.write('B1', 'Length (Inches)', bold)
thermalBatteryWorksheet.write('C1', 'Width (Inches)', bold)
thermalBatteryWorksheet.write('D1', 'Depth (Inches)', bold)
thermalBatteryWorksheet.write('E1', 'Length (mm)', bold)
thermalBatteryWorksheet.write('F1', 'Width (mm)', bold)
thermalBatteryWorksheet.write('G1', 'Depth (mm)', bold)
thermalBatteryWorksheet.write('H1', 'Cost', bold)
thermalBatteryWorksheet.write('I1', 'Quantity', bold)
thermalBatteryWorksheet.write('J1', 'Source', bold)
thermalBatteryWorksheet.write('K1', 'Location', bold)

# Some data we want to write to the thermalBatteryWorksheet.
thermalBatteryMaterials = (
	['EPS Foam', 59.14, 26, 1, 0, 8, '', 'North Air Plenum Back & Front'], 
 	['EPS Foam', 96, 22, 1, 0, 2, '', 'North Air Plenum Mid Shelf'], 
 	['EPS Foam', 22, 17, 1, 0, 2, '', 'North Air Plenum Shelf Support'], 
 	['EPS Foam', 26, 22, 1, 0, 2, '', 'North Air Plenum Ends'], 
 	['EPS Foam', 93.64, 23.906, 2, 0, 2, '', 'North Air Plenum Bottom'], 
 	['EPS Foam', 53.14, 24, 2, 0, 4, '', 'North Air Plenum Top'],
	['EPS Foam', 96, 48, 6, 0, 30, '', 'Thermal Battery Walls'],
	['EPS Foam', 68, 22, 2, 0, 4, '', 'South Plenum Front & Back'],  
	['EPS Foam', 22, 23.906, 2,  0, 4, '', 'South Plenumm Ends'],
	['EPS Foam', 72, 23.906, 2, 0, 4, '', 'South Plenum Top'],
	['EPS Foam', 96, 12, 4, 0, 10, '', 'Side Foundation'],
	['EPS Foam', 96, 36, 4, 0, 4, '', 'End Foundation'],
	['EPS Foam', 36, 36, 4, 0, 2, '', 'End Foundation']


)

# Start from the first cell below the headers.
row = 1
col = 0

# Iterate over the data and write it out row by row.
for material, lengthInch, widthInch, depthInch, cost, quantity, source, location in (thermalBatteryMaterials):
    thermalBatteryWorksheet.write(row, col,     material)
    thermalBatteryWorksheet.write(row, col + 1, lengthInch)
    thermalBatteryWorksheet.write(row, col + 2, widthInch)
    thermalBatteryWorksheet.write(row, col + 3, depthInch)
    thermalBatteryWorksheet.write(row, col + 4, lengthInch * i2m)
    thermalBatteryWorksheet.write(row, col + 5, widthInch * i2m)
    thermalBatteryWorksheet.write(row, col + 6, depthInch * i2m)
    thermalBatteryWorksheet.write(row, col + 7, cost)
    thermalBatteryWorksheet.write(row, col + 8, quantity)
    thermalBatteryWorksheet.write(row, col + 9, source)
    thermalBatteryWorksheet.write(row, col + 10, location)
    print(location)

    row += 1


# Write a total using a formula.
thermalBatteryWorksheet.write(row, 0, 'Total',       bold)
thermalBatteryWorksheet.write(row, 1, '=SUM(B2:B5)', money)


greenhouseWorksheet = workbook.add_worksheet()

# Write some data headers.
greenhouseWorksheet.write('A1', 'Material', bold)
greenhouseWorksheet.write('B1', 'Length (Inches)', bold)
greenhouseWorksheet.write('C1', 'Width (Inches)', bold)
greenhouseWorksheet.write('D1', 'Depth (Inches)', bold)
greenhouseWorksheet.write('E1', 'Length (mm)', bold)
greenhouseWorksheet.write('F1', 'Width (mm)', bold)
greenhouseWorksheet.write('G1', 'Depth (mm)', bold)
greenhouseWorksheet.write('H1', 'Cost', bold)
greenhouseWorksheet.write('I1', 'Quantity', bold)
greenhouseWorksheet.write('J1', 'Source', bold)
greenhouseWorksheet.write('K1', 'Location', bold)

# Some data we want to write to the greenhouseWorksheet.
greenhouseMaterials = (
	['EPS Foam', 96, 48, 6, 0, 17, '', 'Front and Back'],
	['EPS Foam', 71, 42, 2, 0, 6, '', 'North Insulation Ends'],
 	['EPS Foam', 71, 48, 2, 0, 24, '', 'North Insulation Middle'],
 	['EPS Foam', 27.8, 42, 6, 0, 2, '', 'North Insulation Base Ends'], 
 	['EPS Foam', 27.8, 48, 6, 0, 8, '', 'North Insulation Base Middle'], 
 	['EPS Foam', 48.625, 42, 2, 0, 2, '', 'North Insulation Triangle Ends'],
 	['EPS Foam', 48.625, 48, 2, 0, 8, '', 'North Insulation Triangle Middle'], 
 	['EPS Foam', 15.6875, 42, 2, 0, 2, '', 'North Insulation Triangle Ends'],  
 	['EPS Foam', 15.6875, 48, 2, 0, 8, '', 'North Insulation Triangle Middle'], 
 	['EPS Foam', 10.7656, 42, 2, 0, 2, '', 'North Insulation Air Intake Ends'],
 	['EPS Foam', 10.7656, 48, 2, 0, 8, '', 'North Insulation Air Intake Middle'],
 	['EPS Foam', 60.28, 24, 2, 0, 4, '', 'Air Intake'],
 	['EPS Foam', 60.28, 22, 2, 0, 4, '', 'Air Intake'], 
	['EPS Foam', 42, 10, 6, 0, 2, '', 'South PVC Anchor Ends'],
	['EPS Foam', 48, 10, 6, 0, 10, '', 'South PVC Anchor Middle'],
 	


)

# Start from the first cell below the headers.
row = 1
col = 0

# Iterate over the data and write it out row by row.

for material, lengthInch, widthInch, depthInch, cost, quantity, source, location in (greenhouseMaterials):
    greenhouseWorksheet.write(row, col,     material)
    greenhouseWorksheet.write(row, col + 1, lengthInch)
    greenhouseWorksheet.write(row, col + 2, widthInch)
    greenhouseWorksheet.write(row, col + 3, depthInch)
    greenhouseWorksheet.write(row, col + 4, lengthInch * i2m)
    greenhouseWorksheet.write(row, col + 5, widthInch * i2m)
    greenhouseWorksheet.write(row, col + 6, depthInch * i2m)
    greenhouseWorksheet.write(row, col + 7, cost)
    greenhouseWorksheet.write(row, col + 8, quantity)
    greenhouseWorksheet.write(row, col + 9, source)
    greenhouseWorksheet.write(row, col + 10, location)

    row += 1

workbook.close()
