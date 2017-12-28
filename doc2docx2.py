#导入装换的包
import win32com
import xlwt
import os
from win32com.client import Dispatch
#导入界面的包
from doc2docx import Ui_MainWindow
import sys
import doc2docx
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox, QAction

class SimpleDialogForm(Ui_MainWindow, QMainWindow):

    # 文件总数
    totalList = []
    # 转换正确的文件总数
    successList = []
    # 装换错误的文件总数
    errorList = []

    def __init__(self, parent = None):
        super(SimpleDialogForm, self).__init__()
        self.setupUi(self  )  # 在此设置界面

        # 父类的progressBar的值为24,这里设置为0
        self.progressBar.setProperty("value", 0)

        #兴建关于的条目
        self.about = QAction("关于")
        self.contact = QAction("联系")
        #加上帮助菜单栏
        helpMenu = self.menubar.addMenu("帮助")
        #帮助菜单栏上加上条目
        helpMenu.addAction(self.about)
        helpMenu.addAction(self.contact)

        #选中doc文件夹绑定的槽函数
        self.docBtn.clicked.connect(self.setDocUrl)
        #选中docx文件夹绑定的槽函数
        self.docxBtn.clicked.connect(self.setDocxUrl)
        #输入doc路径的输入框文本发生变化时绑定的槽
        self.docLineEdit.textChanged.connect(self.initGUI)
        #输入docx路径的输入框文本发生变化时绑定的槽
        self.docxLineEdit.textChanged.connect(self.initGUI)
        #开始转换按钮绑定的槽函数
        self.startBtn.clicked.connect(self.startConvert)
        #显示关于窗口
        self.about.triggered.connect(self.showAbout)
        #显示联系作者窗口
        self.contact.triggered.connect(self.contactAuthor)

    def writeMsg(self):

        self.textEdit.append("生成（转换信息.xls）")
        docxUrl = self.docxLineEdit.text().strip()
        msgExcel = docxUrl + '/转换信息.xls'
        print (msgExcel)

        excel = xlwt.Workbook(encoding='utf-8')
        # 这个是指定sheet页的名称
        sheet1 = excel.add_sheet('统计信息')
        sheet2 = excel.add_sheet('详细信息')

        sheet1.write(0, 0, '文件总数')
        sheet1.write(0, 1, '转换正确')
        sheet1.write(0, 2, '转换错误')
        sheet1.write(1, 0, len(self.totalList))
        sheet1.write(1, 1, len(self.successList))
        sheet1.write(1, 2, len(self.errorList))

        sheet2.write(0, 0, '文件总数')
        sheet2.write(0, 1, '转换正确')
        sheet2.write(0, 2, '转换错误')

        row = 1
        for x in self.totalList:
            sheet2.write(row, 0, x)
            row += 1

        row = 1
        for x in self.successList:
            sheet2.write(row, 1, x)
            row += 1

        row = 1
        for x in self.errorList:
            sheet2.write(row, 2, x)
            row += 1

        excel.save(msgExcel)
        self.statusbar.showMessage("转换完成，请到生成目录下打开(转换信息.xls)查看详细信息", 10000)
        self.setOp(True)

    def setOp(self, flag):
        self.docLineEdit.setEnabled(flag)
        self.docxLineEdit.setEnabled(flag)
        self.docBtn.setEnabled(flag)
        self.docxBtn.setEnabled(flag)
        self.startBtn.setEnabled(flag)

    def setDocUrl(self):
        #重新选择输入和输出目录时，进度条设置为0，文本框的内容置空
        str = QFileDialog.getExistingDirectory(self,"选中doc文件所在目录",r"C:\Users\Administrator\Desktop")
        self.docLineEdit.setText(str)

    def setDocxUrl(self):
        #重新选择输入和输出目录时，进度条设置为0，文本框的内容置空
        str = QFileDialog.getExistingDirectory(self,"选中生成docx文件所在目录",r"C:\Users\Administrator\Desktop")
        self.docxLineEdit.setText(str)

    #将图形界面的内容各种进度条，文本框初始化
    def initGUI(self):
        self.progressBar.setProperty("value", 0)
        self.textEdit.setText("")

    def initConfig(self):
        # 文件总数
        self.totalList = []
        # 转换正确的文件总数
        self.successList = []
        # 装换错误的文件总数
        self.errorList = []

    def startConvert(self):
        if self.docxLineEdit.text() == '' or self.docxLineEdit.text() == '':
            msgBox = QMessageBox(QMessageBox.Warning, "警告", "请选中要转换的目录")
            msgBox.exec()
            return
        self.setOp(False)
        self.initConfig()
        #获取文本同时去掉空格
        docUrl = self.docLineEdit.text().strip()
        docxUrl = self.docxLineEdit.text().strip()
        word = win32com.client.Dispatch('word.application')
        word.DisplayAlerts = 0
        word.visible = 0
        fileTotal = 0
        total = 0
        for root, dirs, files in os.walk(docUrl):
            fileTotal += len(files)
        for root, dirs, files in os.walk(docUrl):
            self.statusbar.showMessage("正在转换(" + root + ")的文件")
            for name in files:
                length = len(name)
                index = name.rfind('.')
                after = name[index:length]
                after = after.lower()
                total += 1
                if after == '.doc' or after == '.docx':
                    self.totalList.append(name)
                    value = total * 1.0 / fileTotal * 100
                    self.progressBar.setValue(int(value))
                    fileName = root + "/" + name
                    print(fileName)
                    try:
                        doc = word.Documents.Open(fileName)
                        # 这个是保存的目录
                        doc.SaveAs(docxUrl + "/" + fileName.split("/")[-1].split(".")[0] + ".docx", 12)
                        doc.Close()
                        str = name + "转换成功"
                        self.textEdit.append(str)
                        self.successList.append(name)
                    except Exception as e:
                        str = name + "转换失败"
                        self.textEdit.append(str)
                        self.errorList.append(name)
                        continue

        self.writeMsg()

    def showAbout(self):
        msgBox = QMessageBox(QMessageBox.Question, "关于", "选中输入文件夹，输出文件夹，点击开始转换按钮,转换完成\n"
                                                         "后在输出文件夹下有（转换信息.xls）,记录了转换的文件\n"
                                                         "总数,转换正确的文件数,转换错误的文件数等详细信息")
        msgBox.exec()

    def contactAuthor(self):
        msgBox = QMessageBox(QMessageBox.Information, "联系作者", "如您有更好的意见或建议，请联系作者QQ:290059742")
        msgBox.exec()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    main = SimpleDialogForm()
    main.show()
    sys.exit(app.exec_())