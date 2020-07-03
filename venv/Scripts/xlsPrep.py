import openpyxl, datetime

def create_dict(directory, book, sheetTxt):
    """Function to recursively parse the layout tree."""
    myDict = dict()
    pathWs = directory + book
    wb = openpyxl.load_workbook(pathWs)
    sheet = wb[sheetTxt]
    colummn = sheet['A']
    nc = len(colummn)
    for r in range(2, nc + 1, 1):
        c1 = sheet.cell(row=r, column=1).value
        c2 = sheet.cell(row=r, column=2).value
        c1 = str(c1)
        myDict[c1] = c2

    return myDict


def write_doucment(directory, DocNum, CardCode, numatcard, deldate):
    pathWs = directory + 'ORDR - Documents1.xlsx'
    if DocNum == 1:
        pathWs0 = directory + 'ORDR - Documents.xlsx'
        wb = openpyxl.load_workbook(pathWs0)
        wb.save(pathWs)
    wb = openpyxl.load_workbook(pathWs)
    ws = wb['Sheet1']
    colummn = ws['A']
    nc = len(colummn) + 1
    ws.cell(row=nc, column=1).value = DocNum
    ws.cell(row=nc, column=7).value = CardCode
    ws.cell(row=nc, column=6).value = str(deldate.strftime('%Y%m%d'))
    ws.cell(row=nc, column=10).value = numatcard
    wb.save(pathWs)

def write_doucment1(directory, DocNum, CardCode, numatcard, deldate,canceldate,docdueDate):
    pathWs = directory + 'ORDR - DocumentsNew1.xlsx'
    if DocNum == 1:
        pathWs0 = directory + 'ORDR - DocumentsNew.xlsx'
        wb = openpyxl.load_workbook(pathWs0)
        wb.save(pathWs)
    wb = openpyxl.load_workbook(pathWs)
    ws = wb['Sheet1']
    colummn = ws['A']
    nc = len(colummn) + 1

    #docnum
    ws.cell(row=nc, column=1).value = DocNum
    #cardcode
    ws.cell(row=nc, column=8).value = CardCode

    #docDue date
    ws.cell(row=nc, column=7).value = str(docdueDate.strftime('%Y%m%d'))
    # ws.cell(row=nc, column=10).value = numatcard
    ws.cell(row=nc, column=11).value = numatcard
    # ws.cell(row=nc, column=41).value = str(canceldate.strftime('%Y%m%d'))
    ws.cell(row=nc, column=178).value = str(canceldate.strftime('%Y%m%d'))
    #del date
    ws.cell(row=nc, column=177).value = str(deldate.strftime('%Y%m%d'))


    wb.save(pathWs)




def write_line(directory, DocNum, lineno, ItemCode, Qty):
    pathWs = directory + 'RDR1 - Document_Lines1.xlsx'
    if DocNum == 1 and lineno == 0:
        pathWs0 = directory + 'RDR1 - Document_Lines.xlsx'
        wb = openpyxl.load_workbook(pathWs0)
        wb.save(pathWs)
    wb = openpyxl.load_workbook(pathWs)
    ws = wb['Sheet1']
    colummn = ws['A']
    nc = len(colummn) + 1
    ws.cell(row=nc, column=1).value = DocNum
    ws.cell(row=nc, column=2).value = lineno
    ws.cell(row=nc, column=3).value = ItemCode
    ws.cell(row=nc, column=5).value = int(Qty)
    wb.save(pathWs)