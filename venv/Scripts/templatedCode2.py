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

def parse_txtlayout(layout, pge, filename):
    df2 = pd.DataFrame()

    #print("#", count, ",", Text)
    for lt_obj in layout:
        if isinstance(lt_obj, LTAnno) == False:
            Text = lt_obj.get_text().strip()
            Text = Text.replace('\n', ' ')
            x0, y0, x1, y1 = (lt_obj.bbox[0], lt_obj.bbox[1], lt_obj.bbox[2], lt_obj.bbox[3])
            x0, y0, x1, y1 = (round(x0, 4), round(y0, 4), round(x1, 4), round(y1, 4))
            df2 = df2.append({'x0': x0, 'y0': y0, 'x1': x1, 'y1': y1, 'Text': Text, 'pge': pge, 'filename': filename, 'ltType': 'Text'}, ignore_index=True)
            #parse_charlayout(lt_obj, pge, filename)

    return df2

def findTxt(df,text):
    c = df[df['Text'].str.contains(text, regex=False)].iloc[:, [0, 2, 3, 4, 5, 6,7]]
    # Text  ltType pge   x0  x1 y0 y1
    cv = c.values
    #print(cv)
    return cv

def findChild(df,cv,x):
    for i in cv:
        filter1 = df['x1'] >= i[4]
        filter2 = df['pge'] == 1
        filter3 = df['y0'] <= i[5]+x
        filter4 = df['y0'] >= i[5]-x
        filter5 = df['y1'] <= i[6]+x
        filter6 = df['y1'] >= i[6]-x

       # res = df[(filter1 & filter2 & filter3 & filter4 & filter5 & filter6)]
        res = df[(filter1 & filter2 & filter5 & filter6)]
        #res = res['Text'].iloc[0]
        res = res.iloc[:, [0, 2, 3, 4, 5, 6,7]]
        return res

def parseDFT(df,itemDict,vendorDict,DocNum):
    cv = findTxt(df, 'DELIVERY DATE :')
    res = findChild(df, cv, 0)
    deldate = res['Text'].iloc[1]
    deldate = dt.datetime.strptime(deldate, '%m/%d/%Y')

    cv = findTxt(df, 'CANCEL DATE :')
    res = findChild(df, cv, 0)
    cancelDate = res['Text'].iloc[1]
    cancelDate = dt.datetime.strptime( cancelDate , '%m/%d/%Y')

    cv = findTxt(df, 'P.O.NO.:')
    res = findChild(df, cv, 0)
    poN=res['Text'].iloc[1]

    cv = findTxt(df, 'BRANCH')
    res = findChild(df, cv, 0)
    dfcard=res['Text'].iloc[0]
    dfcard = dfcard.lstrip('BRANCH').strip()
    dfcard = dfcard.replace("  ", " ", 1)

    try:
        dfcard = vendorDict[dfcard]
        print(poN,dfcard,DocNum)
        print('numatcard', 'cardcode', 'DocNum', 'deliverydate','cancelDate')
        print(poN, dfcard, DocNum, deldate, cancelDate)
        xlsPrep.write_doucment(directory, DocNum, dfcard, poN, deldate)
        xlsPrep.write_doucment1(directory, DocNum, dfcard, poN, deldate, cancelDate)
        # df2 = df.sort_values(['pge', 'x0', 'y0'], ascending=[True, False, True])
        # print(df.to_string())

    except KeyError:
       print(dfcard ,"not in list")
    return DocNum

#details
def parseDF(df,itemDict,vendorDict,DocNum):
    #print(df.to_string())

    #code to find boundary of column
    c = df[df['Text'].str.contains("IB)", regex=False)].iloc[:, [0, 2, 3, 4]]
    cv = c.values
    print(cv)
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
            #df = my_parse(directory, filename)

            try:
                df = pd.read_pickle(directory + filename + '.pkl')
            except:
                df = my_parse(directory, filename)
            #print(df.to_string())
            DocNum=parseDFT(df, itemDict, vendorDict, count)
            lineNo = 0
            for index, row in df.iterrows():
                Text = row['Text']
                pge = row['pge']

                try:
                    key=Text.split()[0]
                    item = itemDict[Text.split()[0]]
                    cv = findTxt(df, Text.split()[0])
                    res = findChild(df, cv, 0)
                    #print(res)
                    cv1 = findTxt(df, 'IB)')
                    filter1 = res['x0'] >= cv1[0][3]
                    filter3= res['pge'] == 1

                    cv2 = findTxt(df, 'QTY (PC)')
                    filter2 = res['x0'] <= cv2[0][3]
                    dfItem = res[(filter1 &filter2 & filter3)]

                    Qty = float(dfItem.iloc[0][0])
                    #print(key,Qty)
                    print(lineNo, item, Qty, DocNum)
                    xlsPrep.write_line(directory, DocNum, lineNo, item, Qty)
                    lineNo += 1


                    #print(dfItem)
                    #print(Text.split()[0],dfItem['Text'])
                    #print(res['Text'].iloc[1])
                except KeyError:
                    pass


            # cv = findTxt(df, 'IB)')
            # print(cv[0][4])
            # cv = findTxt(df, 'QTY (PC)')
            # print(cv,cv[0][4])

            #parseDF(df, itemDict, vendorDict, count)


            #print(df.to_string())

            #parseDFT(df, itemDict, vendorDict, count)
            #parseDF(df,itemDict,vendorDict,count)



directory = 'metroretailpo/'
# df = my_parse(directory, 'PO 24-19051 DON ANTONIO.pdf')
# cv = findTxt(df, 'DELIVERY DATE')
# print(cv)
# print(df.to_string())
# res = findChild(df, cv, 0)
# # cv = findTxt(df, '01/16/2020')
# # print(cv)
# # res1=findChild(df, cv, .001)
# print(res)

run_batch(directory)
k = input('press enter to exit')
