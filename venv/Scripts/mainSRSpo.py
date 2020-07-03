# Python bytecode 3.5 (3320)
# Embedded file name: C:\Users\Ryan\PycharmProjects\rcsPO\venv\Scripts\main.py
# Decompiled by https://python-decompiler.com
import openpyxl, re, numpy
import pandas as pd
from pdfminer3.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer3.converter import PDFPageAggregator
from pdfminer3.pdfpage import PDFPage
from pdfminer3.layout import LTTextBoxHorizontal, LAParams, LTTextBox, LTTextLine, LTFigure, LTAnno, LTLine, LTRect
from datetime import datetime
from collections import OrderedDict
import xlsPrep
import datetime as dt
from collections import OrderedDict

class myobject:

    def __init__(self, x, y, Text, item, pge):
        self.x = x
        self.y = y
        self.Text = Text
        self.item = item
        self.pge = pge
        self.mylist = list()


directory = 'srspo/'
pathWs = directory + 'SRS BP names.xlsx'
wb = openpyxl.load_workbook(pathWs)
sheet = wb['Sheet1']
colummn = sheet['A']
nc = len(colummn)
vendorDict = dict()
for r in range(2, nc + 1, 1):
    c1 = sheet.cell(row=r, column=1).value
    c2 = sheet.cell(row=r, column=2).value
    c1 = str(c1)
    vendorDict[c1] = c2

pathWs = directory + 'Items revised rcs.xlsx'
wb = openpyxl.load_workbook(pathWs)
sheet = wb['Sheet1']
colummn = sheet['A']
nc = len(colummn)
itemDict = dict()
for r in range(2, nc + 1, 1):
    c1 = sheet.cell(row=r, column=1).value
    c2 = sheet.cell(row=r, column=2).value
    c1 = str(c1)
    itemDict[c1] = c2

def parse_txtlayout(layout, pge, filename):
    df3 = pd.DataFrame()
    for lt_obj in layout:
        if isinstance(lt_obj, LTAnno) == False:
            Text = lt_obj.get_text().strip()
            Text = Text.replace('\n', ' ')
            x0, y0, x1, y1 = (lt_obj.bbox[0], lt_obj.bbox[1], lt_obj.bbox[2], lt_obj.bbox[3])
            x0, y0, x1, y1 = (
             round(x0, 4), round(y0, 4), round(x1, 4), round(y1, 4))
            df3 = df3.append({'x0': x0, 'y0': y0, 'x1': x1, 'y1': y1, 'Text': Text, 'pge': pge, 'filename': filename}, ignore_index=True)
        try:
            item = itemDict[Text]
            yt = str(pge) + ' ' + str(round(y0, 4))
            detailDict[yt] = myobject(x0, y0, Text, item, pge)
        except KeyError:
            pass

    return df3


def parse_page(layout, pge, filename):
    """Function to recursively parse the layout tree."""
    df = pd.DataFrame()
    for lt_obj in layout:
        if isinstance(lt_obj, LTTextBox) or isinstance(lt_obj, LTTextLine):
            df = pd.concat([df, parse_txtlayout(lt_obj, pge, filename)])

    return df


def my_parse(path, filename, DocNum):
    detailDict = OrderedDict()
    document = open(path + filename, 'rb')
    rsrcmgr = PDFResourceManager()
    laparams = LAParams()
    device = PDFPageAggregator(rsrcmgr, laparams=laparams)
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    pge = 0
    df = pd.DataFrame()
    for page in PDFPage.get_pages(document, pagenos=[0, 1]):
        pge = pge + 1
        interpreter.process_page(page)
        layout = device.get_result()
        df = pd.concat([df, parse_page(layout, pge, filename)])

    df = df.sort_values(['pge', 'y0', 'x0'], ascending=[True, False, True])
    df = df.reset_index(drop=True)
    filter1 = df['y0'] == 717.2485
    filter2 = df['x0'] == 39.6854
    filter3 = df['pge'] == 1
    dfadd = df[(filter1 & filter2 & filter3)]
    filter1 = df['y0'] == 702.9073
    filter2 = df['x0'] == 39.6854
    filter3 = df['pge'] == 1
    dfcard = df[(filter1 & filter2 & filter3)]
    filter2 = df['x0'] == 464.8822
    dfpo = df[(filter1 & filter2 & filter3)]
    filter1 = df['y0'] == 647.3624
    filter2 = df['x0'] == 464.8822
    dfdeldate = df[(filter1 & filter2 & filter3)]
    deldate = dfdeldate.iloc[0][0]
    try:
        deldate = dt.datetime.strptime(deldate, '%Y-%m-%d')
    except:
        deldate = dt.datetime.strptime(deldate, '%m/%d/%Y')
    dfadd = dfadd.iloc[0][0]
    poN = dfpo.iloc[0][0]
    try:
        cardfname = vendorDict[dfcard.iloc[0][0]]
    except:
        print(pdf, dfcard.iloc[0][0])
    else:
        xlsPrep.write_doucment(directory, DocNum, cardfname, poN, deldate)
        print('numatcard', 'cardcode', 'DocNum', 'deliverydate', 'address')
        deldate = str(deldate.strftime('%Y-%m-%d'))
        print(poN, cardfname, DocNum, deldate, dfadd)
        print('xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx')
        for index, row in df.iterrows():
            Text = row['Text']
            x0 = row['x0']
            y0 = row['y0']
            pge = row['pge']
            yt = str(pge) + ' ' + str(round(y0, 4))
            try:
                TextS = Text.split()[0]
                try:
                    item = itemDict[TextS]
                    detailDict[yt] = myobject(x0, y0, Text, item, pge)
                except KeyError:
                    try:
                        it = detailDict[yt]
                        TextS = Text.split()
                        for Text in TextS:
                            it.mylist.append(Text)

                    except:
                        pass

            except:
                pass

        count = 0
        print('line', 'item', 'qty', 'page', 'docnum')
        for key, value in detailDict.items():
            Qty = float(value.mylist[1])
            lineNo = count
            xlsPrep.write_line(directory, DocNum, lineNo, value.item, Qty)
            print(lineNo, value.item, Qty, value.pge, DocNum)
            count += 1


def run_batch(directory):
    import os
    count = 0
    for filename in os.listdir(directory):
        if filename.endswith('.PDF') or filename.endswith('.pdf'):
            count = count + 1
            path = os.path.join(directory, filename)
            print('xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx')
            print('processing file :', count, filename)
            my_parse(directory, filename, count)


run_batch(directory)
k = input('press enter to exit')