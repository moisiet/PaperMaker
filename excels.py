from win32com import client
import random
from PyQt5.QtWidgets import QMessageBox
def get_data(path,num):
    data=[]

    try:
        app=client.Dispatch("Excel.Application")
        app.WorkBooks.Open(path)
        app.Visible=False
        ws=app.WorkBooks(1)
    except:
        QMessageBox.warning(None,'错误', path+'数据加载失败！', QMessageBox.Ok)
        print('wrong')
        return data
    sets = set()
    max_row=app.ActiveSheet.UsedRange.Rows.Count
    max_column=app.ActiveSheet.UsedRange.Columns.Count
    print(max_row,max_column)
    while len(sets) < num:
        sets.add(random.randint(2, max_row))
    for i in sets:
        rows = []
        for j in range(1, max_column + 1):
            if j == 1:
                rows.append(remove_order(ws.Sheets(1).Cells(i, j).Value))
            else:
                rows.append(ws.Sheets(1).Cells(i, j).Value)
        data.append(rows)
    print(data)
    return data

# def get_data(path,num):
#     data = []
#     try:
#         wb = excel.load_workbook(path)
#     except:
#         QMessageBox.warning(None,'错误', path+'数据加载失败！', QMessageBox.Ok)
#         return data
#     ws = wb.worksheets[0]
#     sets = set()
#     while len(sets) < num:
#         sets.add(random.randint(2, ws.max_row))
#     for i in sets:
#         rows = []
#         for j in range(1, ws.max_column+1):
#             if j == 1:
#                 rows.append(remove_order(ws.cell(i, j).value))
#             else:
#                 rows.append(ws.cell(i, j).value)
#         data.append(rows)
#     #print(data)
#     return data

def remove_order(content):
    content=content.strip()
    for i in range(len(content)):
        if content[i].isdigit():
            continue
        else:
            if content[i] in [',','，','.',' ','、']:
                return content[i+1:]
            else:
                return content[i:]

def read_xuanze(path,num):
    datas=get_data(path,int(num))
    timu=[]
    daan=[]
    for item in datas:
        timu.append(item[:2])
        daan.append(item[2])
    return timu,daan

def read(path,num):
    datas=get_data(path,int(num))
    timu=[]
    daan=[]
    for item in datas:
        timu.append(item[0])
        daan.append(item[1])
    return timu,daan




if __name__=="__main__":
    get_data(r'D:\PaperMaker\题库\多选题.xlsx',2)
