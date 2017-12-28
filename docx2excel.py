# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'docx2excel.ui'
#
# Created by: PyQt5 UI code generator 5.9
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(500, 430)
        MainWindow.setMinimumSize(QtCore.QSize(500, 430))
        MainWindow.setMaximumSize(QtCore.QSize(500, 430))
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.wordDirLineEdit = QtWidgets.QLineEdit(self.centralwidget)
        self.wordDirLineEdit.setGeometry(QtCore.QRect(60, 30, 260, 20))
        self.wordDirLineEdit.setObjectName("wordDirLineEdit")
        self.wordDirBtn = QtWidgets.QPushButton(self.centralwidget)
        self.wordDirBtn.setGeometry(QtCore.QRect(350, 30, 100, 23))
        self.wordDirBtn.setObjectName("wordDirBtn")
        self.wordTemLineEdit = QtWidgets.QLineEdit(self.centralwidget)
        self.wordTemLineEdit.setGeometry(QtCore.QRect(60, 60, 260, 20))
        self.wordTemLineEdit.setObjectName("wordTemLineEdit")
        self.wordTemBtn = QtWidgets.QPushButton(self.centralwidget)
        self.wordTemBtn.setGeometry(QtCore.QRect(350, 60, 100, 23))
        self.wordTemBtn.setObjectName("wordTemBtn")
        self.excelLineEdit = QtWidgets.QLineEdit(self.centralwidget)
        self.excelLineEdit.setGeometry(QtCore.QRect(60, 90, 260, 20))
        self.excelLineEdit.setObjectName("excelLineEdit")
        self.excelBtn = QtWidgets.QPushButton(self.centralwidget)
        self.excelBtn.setGeometry(QtCore.QRect(350, 90, 100, 23))
        self.excelBtn.setObjectName("excelBtn")
        self.textEdit = QtWidgets.QTextEdit(self.centralwidget)
        self.textEdit.setGeometry(QtCore.QRect(60, 130, 391, 201))
        self.textEdit.setObjectName("textEdit")
        self.progressBar = QtWidgets.QProgressBar(self.centralwidget)
        self.progressBar.setGeometry(QtCore.QRect(60, 350, 270, 23))
        self.progressBar.setProperty("value", 24)
        self.progressBar.setObjectName("progressBar")
        self.startBtn = QtWidgets.QPushButton(self.centralwidget)
        self.startBtn.setGeometry(QtCore.QRect(350, 350, 100, 23))
        self.startBtn.setObjectName("startBtn")
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 500, 23))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "docx2excel"))
        self.wordDirBtn.setText(_translate("MainWindow", "选择Word文件夹"))
        self.wordTemBtn.setText(_translate("MainWindow", "选择Word模板表"))
        self.excelBtn.setText(_translate("MainWindow", "选择Excel表格"))
        self.startBtn.setText(_translate("MainWindow", "开始录入"))

