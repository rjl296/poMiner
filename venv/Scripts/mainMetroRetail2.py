import openpyxl, pandas as pd, xlsPrep, re, numpy
from pdfminer3.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer3.converter import PDFPageAggregator
from pdfminer3.pdfpage import PDFPage
from pdfminer3.layout import LTTextBoxHorizontal, LAParams, LTTextBox, LTTextLine, LTFigure, LTAnno, LTLine, LTRect
import datetime as dt
from collections import OrderedDict
directory = 'metroretailpo/'
path = directory
itemDict = xlsPrep.create_dict(path, 'Metro Retail Items.xlsx', 'Sheet1')
vendorDict = xlsPrep.create_dict(path, 'Metro Retail BP.xlsx', 'Sheet1')

class myobject:

    def __init__(self, x, y, Text, item, pge, x0, x1):
        self.x = x
        self.y = y
        self.Text = Text
        self.item = item
        self.pge = pge
        self.mylist = list()
        self.x0 = x0
        self.x1 = x1


def parse_txtlayout(layout, pge, filename):
    df2 = pd.DataFrame()
    for lt_obj in layout:
        if isinstance(lt_obj, LTAnno) == False:
            Text = lt_obj.get_text().strip()
            Text = Text.replace('\n', ' ')
            x0, y0, x1, y1 = (lt_obj.bbox[0], lt_obj.bbox[1], lt_obj.bbox[2], lt_obj.bbox[3])
            x0, y0, x1, y1 = (
             round(x0, 4), round(y0, 4), round(x1, 4), round(y1, 4))
            df2 = df2.append({'x0': x0, 'y0': y0, 'x1': x1, 'y1': y1, 'Text': Text, 'pge': pge, 'filename': filename}, ignore_index=True)

    return df2


def parse_page(layout, pge, filename):
    """Function to recursively parse the layout tree."""
    df1 = pd.DataFrame()
    for lt_obj in layout:
        if isinstance(lt_obj, LTTextBox) or isinstance(lt_obj, LTTextLine):
            df1 = pd.concat([df1, parse_txtlayout(lt_obj, pge, filename)])

    return df1


def my_parse(path, filename):
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
    df.to_pickle(path + filename + '.pkl')
    return df


def run_batch(directory):
    import os
    count = 0
    for filename in os.listdir(directory):
        if filename.endswith('.PDF') or filename.endswith('.pdf'):
            count = count + 1
            path = os.path.join(directory, filename)
            print('xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx')
            print('processing file :', count, filename)
            try:
                df = pd.read_pickle(directory + filename + '.pkl')
            except:
                df = my_parse(directory, filename)
            else:
                c = df[df['Text'].str.contains('ORDERED')].iloc[:, [2, 3, 4]]
                cv = c.values
                xfilterdict = dict()
                for i in cv:
                    xfilterdict[i[0]] = (i[1], i[2])

                DocNum = count
                filter3 = df['pge'] == 1
                filter1 = df['y0'] == 793.154
                filter2 = df['x0'] == 595.955
                dfpo = df[(filter1 & filter2 & filter3)]
                poN = dfpo.iloc[0][0]

                filter2 = df['x0'] == 148.41
                dfcard = df[(filter2 & filter3)]
                dfcard = dfcard.iloc[0][0]
                dfcard = dfcard.split()
                dfcard = dfcard[0]
                dfcard = vendorDict[dfcard]
                print(poN, dfcard, DocNum)

                filter1 = df['y0'] == 727.634
                filter2 = df['x0'] == 1106.02
                dfdeldate = df[(filter1 & filter2 & filter3)]
                deldate = dfdeldate.iloc[0][0]

                deldate = dt.datetime.strptime(deldate, '%d-%b-%Y')

                print('numatcard', 'cardcode', 'DocNum', 'deliverydate')
                print(poN, dfcard, DocNum, deldate)
                print('xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx')
                xlsPrep.write_doucment(directory, DocNum, dfcard, poN, deldate)
                detailDict = OrderedDict()
                for index, row in df.iterrows():
                    Text = row['Text']
                    x = row['x0']
                    y = row['y0']
                    pge = row['pge']
                    yt = str(pge) + ' ' + str(round(y, 4))
                    try:
                        item = itemDict[Text]
                        x0, x1 = xfilterdict[pge]
                        detailDict[yt] = myobject(x, y, Text, item, pge, x0, x1)
                    except KeyError:
                        pass

                lineNo = 0
                print('line', 'item', 'qty', 'page', 'docnum')
                for key, value in detailDict.items():
                    filter1 = df['y0'] == value.y
                    filter2 = df['x0'] >= value.x0
                    filter3 = df['x1'] <= value.x1
                    dfItem = df[(filter1 & filter2 & filter3)]
                    Qty = float(dfItem.iloc[0][0])
                    xlsPrep.write_line(directory, DocNum, lineNo, value.item, Qty)
                    print(count, value.item, Qty, value.pge, DocNum)
                    lineNo += 1


directory = 'metroretailpo/'
run_batch(directory)
k = input('press enter to exit')