
from PyQt5.Qt import *
from PyQt5.QtCore import *
import sys
import datetime
from excels import *
from words import *
import os

def get_now():
    cur=datetime.datetime.now()
    return "%s-%s-%s"%(cur.year,cur.month,cur.day)

def create_paperdir():
    path=os.getcwd()
    path=os.path.join(path,'试卷')
    if not os.path.exists(path):
        os.makedirs(path)
    return path

class Frame(QDialog):
    def __init__(self,parent):
        super(Frame,self).__init__()
        self.parent=parent
        self.label1=QLabel("版权单位：117营技术保障连",self)
        self.label2=QLabel("策划人：郝洪胜 陈琳",self)
        self.label3=QLabel("作  者：朱陆清",self)
        self.label4=QLabel("日  期：2018年6月2日", self)
        self.label5=QLabel("版权所有 共享使用",self)
        self.label6=QLabel("联系方式：QQ1033857424",self)
        self.mainlayout=QVBoxLayout(self)
        self.mainlayout.addWidget(self.label1)
        self.mainlayout.addWidget(self.label2)
        self.mainlayout.addWidget(self.label3)
        self.mainlayout.addWidget(self.label4)
        self.mainlayout.addWidget(self.label5)
        self.mainlayout.addWidget(self.label6)
        self.setLayout(self.mainlayout)
        self.setWindowFlags(self.windowFlags()|Qt.FramelessWindowHint|Qt.SubWindow)
        self.setStyleSheet("background-color:#90EE90;")
        self.setFixedSize(180, 160)
        self.setFocusPolicy(Qt.ClickFocus)
        self.setFocus(Qt.ActiveWindowFocusReason)
        self.move(parent.pos().x()+(parent.width()-self.width())/2,parent.pos().y()+(parent.height()-self.height())/2)
        self.show()

class Frames(QDialog):
    def __init__(self, parent):
        super(Frames, self).__init__()
        self.parent = parent
        font = QFont()
        font.setPointSize(12)
        self.setFont(font)
        self.label1 = QLabel("使用说明：", self)
        font.setBold(True)
        font.setFamily("黑体")
        font.setPointSize(16)
        self.label1.setFont(font)

        self.label2 = QLabel("1.选择题按照图1模版添加题目，选项和答案", self)
        self.pic1=QLabel('',self)
        path1=os.path.join(os.getcwd(),'res\\select.png')
        self.pic1.setPixmap(QPixmap.fromImage(QImage(path1)))
        self.label4 = QLabel("2.非选择题按照图2模板添加题目和答案", self)
        self.pic2=QLabel('',self)
        img2=QImage(os.path.join(os.getcwd(),'res\\noselect.png'))
        self.pic2.setPixmap(QPixmap.fromImage(img2))
        self.label6 = QLabel("备注：题库采用Excel2007,试卷采用Word2007,其他版本可能有bug，注意使用。", self)

        self.mainlayout = QVBoxLayout(self)
        self.mainlayout.addWidget(self.label1)
        self.mainlayout.addWidget(self.label2)
        self.mainlayout.addWidget(self.pic1)
        self.mainlayout.addWidget(self.label4)
        self.mainlayout.addWidget(self.pic2)
        self.mainlayout.addWidget(self.label6)
        self.setLayout(self.mainlayout)
        self.setWindowFlags(self.windowFlags() | Qt.FramelessWindowHint | Qt.SubWindow)
        self.setStyleSheet("background-color:#90EE90;")
        self.setFixedSize(850, 360)
        self.setFocusPolicy(Qt.ClickFocus)
        self.setFocus(Qt.ActiveWindowFocusReason)
        self.move(parent.pos().x() + (parent.width() - self.width()) / 2,
                  parent.pos().y() + (parent.height() - self.height()) / 2)
        self.show()

    def focusOutEvent(self, event):
        print('aaa')
        self.close()

class Item(QWidget):
    def __init__(self,name,*args,**kwargs):
        super(Item,self).__init__(*args,**kwargs)
        self.setWindowFlags(Qt.FramelessWindowHint)
        self.mainlayout = QHBoxLayout(self)
        self.lTitle = QLabel(name, self)

        font=QFont("黑体")
        font.setPointSize(13)
        self.lTitle.setFont(font)

        if name=="试卷标题：":
            self.eTitle = QLineEdit(self)
            self.mainlayout.addWidget(self.lTitle, 1)
            self.mainlayout.addWidget(self.eTitle, 10)
        else:
            self.lLocation = QLabel("位置", self)
            self.eLocation = QLineEdit(self)
            self.bLocation = QPushButton("...")

            self.lNum = QLabel("数量", self)
            self.eNum = QLineEdit(self)

            self.mainlayout.addWidget(self.lTitle, 1)
            self.mainlayout.addStretch(1)
            self.mainlayout.addWidget(self.lLocation,1)
            self.mainlayout.addWidget(self.eLocation,15)
            self.mainlayout.addWidget(self.bLocation,1)

            self.mainlayout.addStretch(1)
            self.mainlayout.addWidget(self.lNum,1)
            self.mainlayout.addWidget(self.eNum,5)
            self.bLocation.clicked.connect(self.openDlg)

        self.setLayout(self.mainlayout)

    def openDlg(self):
        strFile=QFileDialog.getOpenFileName(self,"选择试卷库",directory=os.getcwd(),filter="DataSource(*.xlsx *.xls)")
        self.eLocation.setText(strFile[0])

class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow,self).__init__()
        self.app,self.doc=None,None
        self.setWindowTitle("试卷助手")
        self.create_menu()

        self.frame=QFrame(self)
        self.setCentralWidget(self.frame)

        self.itemTitle=Item("试卷标题：")
        self.itemBlank=Item("填空题：")
        self.itemSelection=Item("单选题：")
        self.itemMitiSelction=Item('多选题：')

        self.itemJuge=Item("判断题：")
        self.itemQuestion=Item("问答题：")

        self.btn = QPushButton("组 卷", self)
        self.layout = QHBoxLayout()
        self.layout.addStretch(15)
        self.layout.addWidget(self.btn)

        self.mainlayout=QVBoxLayout(self.frame)
        self.mainlayout.addWidget(self.itemTitle)
        self.mainlayout.addWidget(self.itemBlank)
        self.mainlayout.addWidget(self.itemSelection)
        self.mainlayout.addWidget(self.itemMitiSelction)
        self.mainlayout.addWidget(self.itemJuge)
        self.mainlayout.addWidget(self.itemQuestion)
        self.mainlayout.addSpacing(15)
        self.mainlayout.addLayout(self.layout)

        self.setLayout(self.mainlayout)
        self.btn.clicked.connect(self.maker)

    def create_menu(self):
        self.menu=self.menuBar().addMenu("帮助")
        self.about=QAction("作者",self)
        self.help = QAction("使用说明", self)
        self.menu.addAction(self.help)
        self.menu.addAction(self.about)
        self.about.triggered.connect(self.display)
        self.help.triggered.connect(self.display_explain)

    def display(self):
        self.dlg=Frame(self)
    def display_explain(self):
        self.explain=Frames(self)
        
    def get_path(self):
        pathlist=list()
        item=[]
        obj=[self.itemBlank,self.itemSelection,self.itemMitiSelction,self.itemJuge,self.itemQuestion]
        for i in range(len(obj)):
            a=obj[i].eLocation.text().strip()
            b=obj[i].eNum.text().strip()
            if (a=='') or ((not a.endswith('xlsx')) and (not a.endswith('xls'))) or (b=='0') or (b==''):
                continue
            item.append(i+1)
            item.append(a)
            item.append(b)
            pathlist.append(item)
            item=[]
        return pathlist

    def run(self):
        keys=[]
        pathlist=self.get_path()
        for i in range(len(pathlist)):
            if pathlist[i][0]==1:
                timu,daan=read(pathlist[i][1],pathlist[i][2])
                add_tiankong(self.doc,i+1,timu)
            elif pathlist[i][0]==2:
                timu,daan=read_xuanze(pathlist[i][1],pathlist[i][2])
                add_xuanze(self.doc,i+1,timu)
            elif pathlist[i][0]==3:
                timu,daan=read_xuanze(pathlist[i][1],pathlist[i][2])
                add_duoxuan(self.doc,i+1,timu)
            elif pathlist[i][0]==4:
                timu,daan=read(pathlist[i][1],pathlist[i][2])
                add_panduan(self.doc,i+1,timu)
            elif pathlist[i][0]==5:
                timu,daan=read(pathlist[i][1],pathlist[i][2])
                add_wenda(self.doc,i+1,timu)
            keys.append(daan)
        add_heading1(self.doc,'参考答案',fontName='宋体',fontSize=16,Bold=True,fontColor=constants.wdColorBlack,
              alignment=constants.wdAlignParagraphCenter,Indent=0)
        for i in range(len(keys)):
            if pathlist[i][0] == 1:
                add_tiankong_daan(self.doc,i+1,keys[i])
            elif pathlist[i][0]==2:
                add_xuanze_daan(self.doc,i+1,keys[i])
            elif pathlist[i][0]==3:
                add_duoxuan_daan(self.doc,i+1,keys[i])
            elif pathlist[i][0]==4:
                add_panduan_daan(self.doc,i+1,keys[i])
            elif pathlist[i][0]==5:
                add_wenda_daan(self.doc,i+1,keys[i])

    def maker(self):
        self.showMinimized()
        self.app,self.doc=open_document(True)
        add_title(self.doc,self.itemTitle.eTitle.text())
        add_paragraph(self.doc,"单位：             姓名：             分数：            ",
                      fontName='楷体',fontSize=12,Bold=True,fontColor=constants.wdColorBlack,
              alignment=constants.wdAlignParagraphCenter,Indent=0)
        self.run()

        path=os.path.join(create_paperdir(),get_now()+self.itemTitle.eTitle.text()+'.docx')
        self.doc.SaveAs(path)
        # self.doc.Close()
        # self.app.Quit()


if __name__=="__main__":
    app=QApplication(sys.argv)
    item=MainWindow()
    item.show()
    sys.exit(app.exec())