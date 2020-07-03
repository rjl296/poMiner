import openpyxl, pandas as pd, xlsPrep, re, numpy
from pdfminer3.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer3.converter import PDFPageAggregator
from pdfminer3.pdfpage import PDFPage
from pdfminer3.layout import LTTextBoxHorizontal, LAParams, LTTextBox, LTTextLine, LTFigure, LTAnno, LTLine, LTRect,LTChar
import datetime as dt
from collections import OrderedDict
directory = 'mighteemartpo/'
path = directory
itemDict = xlsPrep.create_dict(path, 'Items 10g.xlsx', 'Sheet1')
vendorDict = xlsPrep.create_dict(path, 'Super Grocers BP.xlsx', 'Sheet1')

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

def parse_charlayout(layout, pge, filename):
    #print("#", count, ",", Text)
    for lt_obj in layout:
        if isinstance(lt_obj, LTAnno) == False:
            try:
                x0, y0, x1, y1 = (lt_obj.bbox[0], lt_obj.bbox[1], lt_obj.bbox[2], lt_obj.bbox[3])
                x0, y0, x1, y1 = (round(x0, 4), round(y0, 4), round(x1, 4), round(y1, 4))
            except:
                x0, y0, x1, y1 =55,55,55,55
            print(lt_obj.get_text(),x0,y0,x1,y1)
        #Text = lt_obj.get_text().strip()
        #Text = Text.replace('\n', ' ')

def parse_txtlayout(layout, pge, filename):
    df2 = pd.DataFrame()

    #print("#", count, ",", Text)
    for lt_obj in layout:

        if isinstance(lt_obj, (LTChar, LTAnno)):
            parse_charlayout(t_obj, pge, filename)
        if isinstance(lt_obj, LTAnno) == False:
            Text = lt_obj.get_text().strip()
            Text = Text.replace('\n', ' ')
            x0, y0, x1, y1 = (lt_obj.bbox[0], lt_obj.bbox[1], lt_obj.bbox[2], lt_obj.bbox[3])
            x0, y0, x1, y1 = (round(x0, 4), round(y0, 4), round(x1, 4), round(y1, 4))
            df2 = df2.append({'x0': x0, 'y0': y0, 'x1': x1, 'y1': y1, 'Text': Text, 'pge': pge, 'filename': filename, 'ltType': 'Text'}, ignore_index=True)
            #parse_charlayout(lt_obj, pge, filename)

    return df2

def parse_line(lt_obj, pge, filename):
    dfLine = pd.DataFrame()

    Text = ""
    x0, y0, x1, y1 = (lt_obj.bbox[0], lt_obj.bbox[1], lt_obj.bbox[2], lt_obj.bbox[3])
    x0, y0, x1, y1 = (round(x0, 4), round(y0, 4), round(x1, 4), round(y1, 4))
    dfLine = dfLine.append({'x0': x0, 'y0': y0, 'x1': x1, 'y1': y1, 'Text': Text, 'pge': pge, 'filename': filename, 'ltType': 'Line'}, ignore_index=True)

    return dfLine

def parse_page(layout, pge, filename):

    """Function to recursively parse the layout tree."""
    df1 = pd.DataFrame()
    #print('left, top=, width= ,height=,')
    for lt_obj in layout:
        if isinstance(lt_obj, LTTextBox) or isinstance(lt_obj, LTTextLine):
            #Text = lt_obj.get_text().strip()
           # Text = Text.replace('\n', ' ')
            #print("#",count,",",Text)
            df1 = pd.concat([df1, parse_txtlayout(lt_obj, pge, filename)])

        # if isinstance(lt_obj,LTLine):
        #
        #     a =[ lt_obj.bbox[0], lt_obj.bbox[1], lt_obj.bbox[2], lt_obj.bbox[3]]
        #    # print("LTLINE",x0, y0, x1, y1)
        #     #print ('left="%s" top="%s" width="%s" height="%s"' % (int(a[0]), int(a[1]), int(a[2]-a[0]), int(a[3]-a[1])))
        #     #x0, y0, x1, y1 = (lt_obj.bbox[0], lt_obj.bbox[1], lt_obj.bbox[2], lt_obj.bbox[3])
        #
        #     #print(a[0], a[1], a[2] - a[0], a[3] - a[1])
        #     df1 = pd.concat([df1, parse_line(lt_obj, pge, filename)])

            #print("LTLINE",x0, y0, x1, y1)

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

def parseDFT(df,itemDict,vendorDict,DocNum):
    #
    # c = df[df['Text'].str.contains('DELIVERY DATE :')].iloc[:, [0, 2, 3, 4,5,6]]
    # #print(df.to_string())
    # #print(c)
    # cv=c.values
    # print(cv)
    c = df[df['Text'].str.contains('DELIVERY DATE :')].iloc[:, [0, 2, 3, 4, 5, 6]]
    cv = c.values
    for i in cv:
        filter1 = df['x0'] > i[2]
        filter2 = df['y1'] == i[5]
        filter3 = df['pge'] == 1
        deldate = df[(filter1 & filter2 & filter3)]
        deldate = deldate['Text'].iloc[0]
        deldate = dt.datetime.strptime(deldate, '%m/%d/%Y')

    c = df[df['Text'].str.contains('P.O.NO.:')].iloc[:, [0, 2, 3, 4, 5, 6]]
    cv = c.values
    for i in cv:
        filter1 = df['x0'] > i[2]
        filter2 = df['y1'] == i[5]
        filter3 = df['pge'] == 1
        poN = df[(filter1 & filter2 & filter3)]
        poN=poN['Text'].iloc[0]

    c = df[df['Text'].str.contains('BRANCH')].iloc[:, [0, 2, 3, 4, 5, 6]]
    dfcard =c.values[0][0]
    dfcard = dfcard.lstrip('BRANCH').strip()
    dfcard = dfcard.replace("  ", " ", 1)
    try:
        dfcard = vendorDict[dfcard]
        print(poN,dfcard,DocNum)
        print('numatcard', 'cardcode', 'DocNum', 'deliverydate')
        print(poN, dfcard, DocNum, deldate)
        xlsPrep.write_doucment(directory, DocNum, dfcard, poN, deldate)

        # df2 = df.sort_values(['pge', 'x0', 'y0'], ascending=[True, False, True])
        # print(df.to_string())
    except KeyError:
       print(dfcard ,"not in list")



    #print(poN,dfcard,DocNum)
    #print('numatcard', 'cardcode', 'DocNum', 'deliverydate')
    #print(poN, dfcard, DocNum, deldate)
    #xlsPrep.write_doucment(directory, DocNum, dfcard, poN, deldate)

    #df2 = df.sort_values(['pge', 'x0', 'y0'], ascending=[True, False, True])
    #print(df.to_string())

#details
def parseDF(df,itemDict,vendorDict,DocNum):
    #print(df.to_string())

    #code to find boundary of column
    c = df[df['Text'].str.contains("IB)", regex=False)].iloc[:, [0, 2, 3, 4]]
    cv = c.values
    #print(cv)
    xfilterdict = dict()
    for i in cv:
        xfilterdict[i[1]] = (i[2], i[3])

    c = df[df['Text'].str.contains("QTY (PC)", regex=False)].iloc[:, [0, 2, 3, 4]]
    cv = c.values
   # print(cv)
    xfilterdict2 = dict()
    for i in cv:
        xfilterdict2[i[1]] = (i[2], i[3])

    detailDict = OrderedDict()

    for index, row in df.iterrows():
        Text = row['Text']
        x = row['x0']
        y = row['y0']
        pge = row['pge']
        yt = str(pge) + ' ' + str(round(y, 4))

        try:

            item = itemDict[Text.split()[0]]
            #x0, x1 = xfilterdict[pge]
            x0, _ = xfilterdict[pge]
            _, x1 = xfilterdict2[pge]
            detailDict[yt] = myobject(x, y, Text, item, pge, x0, x1)
            #print(Text.split()[0],item)
        except KeyError:
            pass

    lineNo = 0
    # for key, value in detailDict.items():
    #     print(value.y)
    print('line', 'item', 'qty', 'page', 'docnum')
    for key, value in detailDict.items():
        filter1 = df['y0'] == value.y
        filter2 = df['x0'] >=value.x0
        filter3 = df['x1'] <=value.x1
        dfItem = df[(filter1 & filter2 & filter3)]
       # dfItem = df[(filter1 )]
        #print(dfItem.to_string())
        Qty = float(dfItem.iloc[0][0])
        #print(value.item, Qty)
        xlsPrep.write_line(directory, DocNum, lineNo, value.item, Qty)
        print(lineNo, value.item, Qty, value.pge, DocNum)
        lineNo += 1


def run_batch(directory):
    import os
    count = 0
    for filename in os.listdir(directory):
        if filename.endswith('.PDF') or filename.endswith('.pdf'):
            count = count + 1
            path = os.path.join(directory, filename)

            print('xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx')
            print('processing file :#',count,"-", filename)
            df = my_parse(directory, filename)
            #print(df.to_string())
            # try:
            #     df = pd.read_pickle(directory + filename + '.pkl')
            # except:
            #     df = my_parse(directory, filename)
            # #print(df.to_string())

            #parseDFT(df, itemDict, vendorDict, count)
            #parseDF(df,itemDict,vendorDict,count)



directory = 'mighteemartpo/'
run_batch(directory)
k = input('press enter to exit')
