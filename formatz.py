from openpyxl.utils import get_column_letter as getColLetters
from openpyxl.styles import Alignment
from openpyxl.worksheet.table import Table


def Alt_HOI(worksheet):
    column_widths = {}
    for row in worksheet.iter_rows():
        for cell in row:
            if cell.value:
                # Calculate the maximum length of each column
                column_widths[cell.column_letter] = max(
                    (column_widths.get(cell.column_letter, 0)),
                    len(str(cell.value))
                )
    # Add a little extra width for aesthetics
    for col, width in column_widths.items():
        worksheet.column_dimensions[col].width = width + 3


def Alt_A(worksheet):
    max_row = worksheet.max_row
    max_col = worksheet.max_column
    return f"A1:{getColLetters(max_col)}{max_row}"


def tableFormat(wb):
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        ws.freeze_panes = 'A2'
        for col in ws.iter_cols(max_row=1, max_col=ws.max_column):
            for cell in col:
                cell.alignment = Alignment(horizontal='left')
        ws.sheet_view.showGridLines = False
        wsData = Alt_A(ws)
        ws.add_table(Table(displayName=f'{sheet}', ref=wsData))
        ws = Alt_HOI(ws)
