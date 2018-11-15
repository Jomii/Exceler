import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from sys import argv
import re
from collections import defaultdict

# TODO: Add errorchecking for filename input
# Make directory independent file input.
# Better way of accessing input file (instead of argv)

tulot = []
menot = []
dataTuple = [] # Contains date, transaction and source of transaction.

# If no command-line parameter for filename, then close the program.
if len(argv) < 2:
    exit()

transactionFilename = argv[1]

with open(transactionFilename) as f:
    next(f)
    next(f)
    next(f)
    for line in f:
        parts = line.strip('\n').split('\t')

        if len(parts) > 1:
            t = int(parts[2].split('.')[0]), float(parts[3].replace(',', '.')), parts[4]

            if t[1] > 0:
                tulot.append(t)
            else:
                menot.append(t)

            dataTuple.append(t)

# Excel filename where the data gets appended to.
fileName = 'test.xlsx'
print(f"Adding data from {transactionFilename} to {fileName}")

# Open workbook and select the second to last sheet.
wb = openpyxl.load_workbook(fileName)
ws = wb.create_sheet("Kuukausi", -1)

""" Add income data to cells. """
ws["A1"] = "Pvm."
ws["B1"] = "€"
ws["C1"] = "Saaja/Maksaja"

i = 2

for paiva in tulot:
    ws[f"A{i}"] = paiva[0]
    ws[f"B{i}"] = paiva[1]
    ws[f"C{i}"] = paiva[2]

    i = i + 1

i = i + 2 # Separate values by two rows.

""" Add expense data to cells. """
ws[f"A{i - 1}"] = "Pvm."
ws[f"B{i - 1}"] = "€"
ws[f"C{i - 1}"] = "Saaja/Maksaja"

for paiva in menot:
    ws[f"A{i}"] = paiva[0]
    ws[f"B{i}"] = paiva[1]
    ws[f"C{i}"] = paiva[2]

    i = i + 1


ws.insert_cols(1, amount=1) # Add a column above the data in the sheet.
ws.insert_rows(1, amount=1) # Add a row on the left of the data.
ws.column_dimensions["C"].width = 10
ws.column_dimensions["D"].width = 35

ws["A1"] = "TULOT"
ws[f"A{len(tulot) + 3}"] = "MENOT"

tulotC = ws["E1"]
menotC = ws[f"E{len(tulot) + 3}"]
g1 = ws["G1"]
h1 = ws["H1"]

# Add the balance of the month to a cell.
g1.value = "Balanssi:"
h1.value = f"=SUM(E1:E{len(tulot) + 3})"

# Apply styles to a range of cells as if they were a single cell.
def style_range(ws, cell_range, fill=None, font=None, alignment=None, merge=False):
    # Merge cells in range and apply style to them.
    if merge:
        first_cell = ws[cell_range.split(":")[0]]
        ws.merge_cells(cell_range)
        first_cell.fill = fill
        first_cell.alignment = alignment
        first_cell.font = font
    # Apply styles to each cell individually.
    else:
        rows = ws[cell_range]
        for row in rows:
            for c in row:
                if fill:
                    c.fill = fill
                if font:
                    c.font = font
                if alignment:
                    c.alignment = alignment


# Apply styling to the cells.
greenColor = PatternFill(fill_type='solid', start_color='c6e0b4', end_color='c6e0b4')
redColor = PatternFill(fill_type='solid', start_color='fce4d6', end_color='fce4d6')
yellowColor = PatternFill(fill_type='solid', start_color='ffe699', end_color='ffe699')
al = Alignment(horizontal="center", vertical="center")
al2 = Alignment(horizontal="center", vertical="center", text_rotation=90)
font = Font(bold=True)
double = Side(border_style="medium")
border = Border(top=double, left=double, right=double, bottom=double)
# Add colors and cell alignments
style_range(ws, 'A1:D2', fill=greenColor)
style_range(ws, f'A{len(tulot) + 3}:D{len(tulot) + 4}', fill=redColor)
style_range(ws, f'A1:A{len(tulot) + 2}', fill=greenColor, font=font, alignment=al2, merge=True)
style_range(ws, f'A{len(tulot) + 3}:A{i}', fill=redColor, font=font, alignment=al2, merge=True)
# For balance
style_range(ws, 'G1:H1', font=font, fill=yellowColor)
# Add fonts
style_range(ws, 'B2:D2', font=font)
style_range(ws, f'B{len(tulot) + 4}:D{len(tulot) + 4}', font=font)
# Center Date and currency
style_range(ws, f'B2:C{i}', alignment=al)
# Merge, center and add borders for totals
style_range(ws, 'E1:E2', fill=greenColor, font=font, alignment=al, merge=True)
ws["E1"].border = border
ws["E2"].border = border
style_range(ws, f'E{len(tulot) + 3}:E{len(tulot) + 4}', fill=redColor, font=font, alignment=al, merge=True)
ws[f"E{len(tulot) + 3}"].border = border
ws[f"E{len(tulot) + 4}"].border = border
tulotC.value = sum(map(lambda x: x[1], tulot))
menotC.value = sum(map(lambda x: x[1], menot))

wb.save(fileName) # Save changes to the workbook.