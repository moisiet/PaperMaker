#from win32com.client.constant import constants
from win32com.gen_py.wordconsts import constants
from win32com import client

def number_to_character(num):
    num=str(num)
    dic={'1':'一','2':'二','3':'三','4':'四','5':'五','6':'六','7':'七','8':'八','9':'九'}
    return dic.get(num,'0')

def add_content_xuanze(doc,content,spacenums=0):
    if not content:
        return
    lens = len(content)
    for i in range(lens):
        add_paragraph(doc, str(i + 1) + '、' + content[i][0])
        add_paragraph(doc,content[i][1])
        add_spacing(doc,spacenums)

def add_content(doc,content,spacenums=0):
    if not content:
        return
    lens = len(content)
    for i in range(lens):
        add_paragraph(doc, str(i + 1) + '、' + content[i])
        add_spacing(doc,spacenums)


def open_document(visible=False):
    app=client.Dispatch("Word.Application")
    app.Visible=visible
    app.DisplayAlerts = 0
    doc=app.Documents.Add()
    return app,doc

def add_title(doc,content,fontName='宋体',fontSize=16,Bold=True,
              fontColor=constants.wdColorBlack,
              alignment=constants.wdAlignParagraphCenter,
              Indent=0
              ):
    doc.Content.InsertAfter(content)
    range=doc.Paragraphs.Last.Range
    range.Font.Name=fontName
    range.Font.Size=fontSize
    range.Font.Bold=Bold
    range.Font.Color=fontColor
    range.ParagraphFormat.Alignment =alignment
    range.ParagraphFormat.FirstLineIndent=Indent

def add_heading1(doc,content,fontName='黑体',fontSize=12,Bold=False,
              fontColor=constants.wdColorBlack,
              alignment=constants.wdAlignParagraphLeft,
              Indent=0
              ):
    doc.Content.InsertParagraphAfter()
    add_title(doc, content, fontName, fontSize, Bold, fontColor, alignment, Indent)

def add_paragraph(doc,content,fontName='宋体',fontSize=10,Bold=False,
                  fontColor=constants.wdColorBlack,
                  alignment=constants.wdAlignParagraphLeft,
                  Indent=0):
    doc.Content.InsertParagraphAfter()
    add_title(doc,content,fontName,fontSize,Bold,fontColor,alignment,Indent)


def add_spacing(doc,count=1):
    for i in range(count):
        doc.Content.InsertParagraphAfter()

def add_ABCD(doc,content):
    if not content:
        return
    doc.Content.InsertParagraphAfter()
    lens=len(content)
    for i in range(lens):
        if i%5==0:
            if lens-i-1>=5:
                add_title(doc,'{0}-{1}'.format(i+1,i+5),fontName='宋体',fontSize=10,Bold=False,
                          fontColor=constants.wdColorBlack,alignment=constants.wdAlignParagraphLeft,Indent=0)
                #doc.Paragraphs.Last.Range.InsertAfter('{0}-{1}'.format(i+1,i+5))
            else:
                add_title(doc, '{0}-{1}'.format(i+1,lens), fontName='宋体', fontSize=10, Bold=False,
                          fontColor=constants.wdColorBlack, alignment=constants.wdAlignParagraphLeft, Indent=0)
                #doc.Paragraphs.Last.Range.InsertAfter('{0}-{1}'.format(i+1,lens))
        if i%5==4:
            add_title(doc, content[i]+'   ', fontName='宋体', fontSize=10, Bold=False,
                      fontColor=constants.wdColorBlack, alignment=constants.wdAlignParagraphLeft, Indent=0)
            #doc.Paragraphs.Last.Range.InsertAfter(content[i]+'   ')
        elif i==lens-1:
            add_title(doc, content[i], fontName='宋体', fontSize=10, Bold=False,
                      fontColor=constants.wdColorBlack, alignment=constants.wdAlignParagraphLeft, Indent=0)
            #doc.Paragraphs.Last.Range.InsertAfter(content[i])
        else:
            add_title(doc, content[i] + ',', fontName='宋体', fontSize=10, Bold=False,
                      fontColor=constants.wdColorBlack, alignment=constants.wdAlignParagraphLeft, Indent=0)
            #doc.Paragraphs.Last.Range.InsertAfter(str(content[i]) + ',')

def add_xuanze_daan(doc,num,content):
    if not content:
        return
    add_heading1(doc,number_to_character(num)+'.单选题：')
    add_ABCD(doc,content)

def add_xuanze(doc,num,content):
    if not content:
        return
    add_heading1(doc, number_to_character(num) + '.单选题（每题 分）：')
    add_content_xuanze(doc,content)

def add_duoxuan_daan(doc,num,content):
    if not content:
        return
    add_heading1(doc,number_to_character(num)+'.多选题：')
    add_ABCD(doc,content)

def add_duoxuan(doc,num,content):
    if not content:
        return
    add_heading1(doc, number_to_character(num) + '.多选题（每题 分）：')
    add_content_xuanze(doc,content)

def add_tiankong_daan(doc,num,content):
    if not content:
        return
    add_heading1(doc, number_to_character(num) + '.填空题：')
    add_content(doc,content)

def add_tiankong(doc,num,content):
    if not content:
        return
    add_heading1(doc, number_to_character(num) + '.填空题（每题 分）：')
    add_content(doc, content)

def add_panduan_daan(doc,num,content):
    if not content:
        return
    add_heading1(doc, number_to_character(num) + '.判断题：')
    add_ABCD(doc, content)

def add_panduan(doc, num, content):
    if not content:
        return
    add_heading1(doc, number_to_character(num) + '.判断题（每题 分）：')
    add_content(doc,content)

def add_wenda_daan(doc,num,content):
    if not content:
        return
    add_heading1(doc, number_to_character(num) + '.问答题：')
    add_content(doc,content)

def add_wenda(doc,num,content,spacenums=3):
    if not content:
        return
    add_heading1(doc, number_to_character(num) + '.问答题（每题 分）：')
    add_content(doc,content,spacenums)






