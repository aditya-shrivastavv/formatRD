import os ,openpyxl
from xls2xlsx import XLS2XLSX


COLS_TO_DEL = [20,19,18,17,16] # must be in descending order
LOGO_DIMENSIONS = (151.18, 64.62) # width, height


def delete_columns(sheet):
    for column in COLS_TO_DEL:
        sheet.delete_cols(column)

def corrections(sheet):
    # Last Created Date & Time column name
    sheet.unmerge_cells(
        start_row = 13,
        start_column = 17,
        end_row = 13,
        end_column = 18
    )
    c = sheet['R13']
    c.value = 'Last Created Date & Time'
    c.alignment = openpyxl.styles.Alignment(vertical='center')
    c.alignment = openpyxl.styles.Alignment(wrap_text=True)
    c.fill = openpyxl.styles.PatternFill(
        fill_type='solid',
        start_color='ccccff',
        end_color='ccccff'
    )

    sheet.unmerge_cells(
        start_row = 13,
        start_column = 19,
        end_row = 13,
        end_column = 20
    )

    # row 2 and 3
    row23 = [2,3]
    for row in row23:
        sheet.unmerge_cells(
            start_row = row,
            start_column = 2,
            end_row = row,
            end_column = 19
        )
        if row == 2:
            c = sheet['B2']
            c.value = 'RECURRING DEPOSIT INSTALLMENT REPORT'
            c.alignment = openpyxl.styles.Alignment(horizontal='center')
        elif row == 3:
            c = sheet['B3']
            c.value = 'Search Criteria'
            c.alignment = openpyxl.styles.Alignment(horizontal='left')
        sheet.merge_cells(
            start_row = row,
            start_column = 2,
            end_row = row,
            end_column = 18
        )

    # row 4, 5, 6, 7, 9
    row45679 = [4,5,6,7,9]
    for row in row45679:
        val = sheet['I' + str(row)].value
        sheet.merged_cells.remove('I' + str(row) + ':X' + str(row))
        sheet['I' + str(row)].value = val
        sheet.merge_cells(
            start_row = row,
            start_column = 9,
            end_row = row,
            end_column = 18
        )

    # row 10 "search results"
    sheet.merged_cells.remove('B10:Y10')
    sheet['B10'].value = "Search Results"
    sheet.merge_cells(
        start_row = 10,
        start_column = 2,
        end_row = 10,
        end_column = 18
    )

    # delete column (all ending ones)
    sheet.delete_cols(19,10)


def addLogo(sheet):
    logo = openpyxl.drawing.image.Image('logo.png')
    logo.width = LOGO_DIMENSIONS[0]
    logo.height = LOGO_DIMENSIONS[1]
    sheet.add_image(logo, 'E1')


def resizeRows(sheet):
    row = 1
    sheet.row_dimensions[row].height = 52
    row = 2

    # doc data
    for _row in range(row, 13):
        sheet.row_dimensions[_row].height = 15

    row = 13

    # main table
    for _row in range(row, sheet.max_row + 1):
        sheet.row_dimensions[_row].height = 30


def resizeColumns(sheet):
    sheet.column_dimensions['E'].width = 11.9
    sheet.column_dimensions['F'].width = 15.4
    sheet.column_dimensions['G'].width = 21
    sheet.column_dimensions['P'].width = 14.8
    sheet.column_dimensions['R'].width = 14.5


def fontSize12ndColorBlack(sheet):
    for i in range(1, sheet.max_row + 1):
        for cell in sheet[i]:
            cell.font = openpyxl.styles.Font(size=12, color='000000', bold=True)


def applyBorder(sheet):
    start = 13
    for _row in range(start, sheet.max_row + 1):
        row = sheet['C' + str(_row):'R' + str(_row)]
        for cell in row[0]:
            cell.border = openpyxl.styles.borders.Border(
                left=openpyxl.styles.borders.Side(border_style='thin'),
                right=openpyxl.styles.borders.Side(border_style='thin'),
                top=openpyxl.styles.borders.Side(border_style='thin'),
                bottom=openpyxl.styles.borders.Side(border_style='thin')
            )
                          

def printSetup(sheet):
    sheet.print_area = f'A1:R{sheet.max_row}'

    sheet.page_margins.left = 0.25
    sheet.page_margins.right = 0.25
    sheet.page_margins.top = 0.75
    sheet.page_margins.bottom = 0.75

    sheet.page_setup.fitToPage = True


def increaseFontSizeForRow(sheet, row, font_size):
    for cell in sheet[row]:
        cell.font = openpyxl.styles.Font(size=font_size)
    
    sheet.row_dimensions[row].height = 24



def formatTheFile(file: str):
    spreadsheet = openpyxl.load_workbook(file)
    sheet = spreadsheet.active

    # 1. delete unwanted columns
    delete_columns(sheet)

    # 2. minor corrections
    corrections(sheet)

    # 2. add logo
    addLogo(sheet)

    # 3. resize some columns
    resizeColumns(sheet)

    # 4. resize rows
    resizeRows(sheet)

    # 5. change font size
    fontSize12ndColorBlack(sheet)

    # 6. apply border   
    applyBorder(sheet)

    # 7. highlight reference id row
    increaseFontSizeForRow(sheet, 6, 18)

    # 8. set print area
    printSetup(sheet)

    spreadsheet.save(file)



def convertFiles() -> int:
    allFiles = os.listdir()
    xlsFiles = set()
    
    # get desired files from all files
    for file in allFiles:
        if file.endswith('.xls'):
            xlsFiles.add(file)

    numberOfFiles = len(xlsFiles)
    if numberOfFiles == 0:
        raise Exception("There are no files with extension '.xls in current folder")

    for file in xlsFiles:
        convert = XLS2XLSX(file)
        newFile = file.split('.')[0] + '.xlsx'
        convert.to_xlsx(newFile)

        formatTheFile(newFile)

        # remove xls file
        os.remove(file)

    return numberOfFiles



def verifyNumberOfFiles(initialNumber):
    allFiles = os.listdir()
    xlsFiles = set()

    # get desired files from all files
    for file in allFiles:
        if file.endswith('.xlsx'):
            xlsFiles.add(file)
    
    # verify number of files
    if initialNumber != len(xlsFiles):
        raise Exception("Number of files are not equal it means some files are skipped")



def run():
    initialNumber = convertFiles()

    verifyNumberOfFiles(initialNumber)



if __name__ == "__main__":
    run()
    print('Dear. Paritosh Ji, All files are formatted successfully!')
