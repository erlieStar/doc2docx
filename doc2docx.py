# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'doc2docx.ui'
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
        self.docLineEdit = QtWidgets.QLineEdit(self.centralwidget)
        self.docLineEdit.setGeometry(QtCore.QRect(60, 40, 260, 20))
        self.docLineEdit.setObjectName("docLineEdit")
        self.docBtn = QtWidgets.QPushButton(self.centralwidget)
        self.docBtn.setGeometry(QtCore.QRect(340, 40, 100, 23))
        self.docBtn.setObjectName("docBtn")
        self.progressBar = QtWidgets.QProgressBar(self.centralwidget)
        self.progressBar.setGeometry(QtCore.QRect(60, 350, 270, 23))
        self.progressBar.setProperty("value", 24)
        self.progressBar.setObjectName("progressBar")
        self.docxBtn = QtWidgets.QPushButton(self.centralwidget)
        self.docxBtn.setGeometry(QtCore.QRect(340, 80, 100, 23))
        self.docxBtn.setObjectName("docxBtn")
        self.docxLineEdit = QtWidgets.QLineEdit(self.centralwidget)
        self.docxLineEdit.setGeometry(QtCore.QRect(60, 80, 260, 20))
        self.docxLineEdit.setObjectName("docxLineEdit")
        self.textEdit = QtWidgets.QTextEdit(self.centralwidget)
        self.textEdit.setGeometry(QtCore.QRect(60, 140, 380, 190))
        self.textEdit.setObjectName("textEdit")
        self.startBtn = QtWidgets.QPushButton(self.centralwidget)
        self.startBtn.setGeometry(QtCore.QRect(340, 350, 100, 23))
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
        MainWindow.setWindowTitle(_translate("MainWindow", "doc2docx"))
        self.docBtn.setText(_translate("MainWindow", "选择输入文件夹"))
        self.docxBtn.setText(_translate("MainWindow", "选择输出文件夹"))
        self.startBtn.setText(_translate("MainWindow", "开始转换"))

