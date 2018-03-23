# -*- coding: utf-8 -*-

import sys
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QMainWindow, QApplication,QFileDialog

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        #MainWindow.resize(425, 477)
        MainWindow.setFixedSize(425, 477)
        MainWindow.setStyleSheet("QMainWindow {background: 'white';}");
        MainWindow.setWindowIcon(QtGui.QIcon("kml_icon.png"));
        self.centralWidget = QtWidgets.QWidget(MainWindow)
        self.centralWidget.setObjectName("centralWidget")
        self.verticalLayoutWidget = QtWidgets.QWidget(self.centralWidget)
        self.verticalLayoutWidget.setGeometry(QtCore.QRect(30, 10, 371, 61))
        self.verticalLayoutWidget.setObjectName("verticalLayoutWidget")
        self.verticalLayout_2 = QtWidgets.QVBoxLayout(self.verticalLayoutWidget)
        self.verticalLayout_2.setContentsMargins(11, 11, 11, 11)
        self.verticalLayout_2.setSpacing(6)
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.label_5 = QtWidgets.QLabel(self.verticalLayoutWidget)
        self.label_5.setObjectName("label_5")
        self.verticalLayout_2.addWidget(self.label_5)
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_2.setSpacing(6)
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.lineEdit = QtWidgets.QLineEdit(self.verticalLayoutWidget)
        self.lineEdit.setObjectName("lineEdit")
        self.horizontalLayout_2.addWidget(self.lineEdit)
        self.selectExcelButton = QtWidgets.QToolButton(self.verticalLayoutWidget)
        self.selectExcelButton.setObjectName("selectExcelButton")

        self.selectExcelButton.clicked.connect(self.excelDialog)

        self.horizontalLayout_2.addWidget(self.selectExcelButton)
        self.verticalLayout_2.addLayout(self.horizontalLayout_2)
        self.verticalLayoutWidget_2 = QtWidgets.QWidget(self.centralWidget)
        self.verticalLayoutWidget_2.setGeometry(QtCore.QRect(30, 90, 191, 181))
        self.verticalLayoutWidget_2.setObjectName("verticalLayoutWidget_2")
        self.verticalLayout_3 = QtWidgets.QVBoxLayout(self.verticalLayoutWidget_2)
        self.verticalLayout_3.setContentsMargins(11, 11, 11, 11)
        self.verticalLayout_3.setSpacing(6)
        self.verticalLayout_3.setObjectName("verticalLayout_3")
        self.horizontalLayout_7 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_7.setSpacing(6)
        self.horizontalLayout_7.setObjectName("horizontalLayout_7")
        self.label = QtWidgets.QLabel(self.verticalLayoutWidget_2)
        self.label.setObjectName("label")
        self.horizontalLayout_7.addWidget(self.label)
        self.nameTextBox = QtWidgets.QLineEdit(self.verticalLayoutWidget_2)
        self.nameTextBox.setObjectName("nameTextBox")
        self.horizontalLayout_7.addWidget(self.nameTextBox)
        self.verticalLayout_3.addLayout(self.horizontalLayout_7)
        self.horizontalLayout_9 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_9.setSpacing(6)
        self.horizontalLayout_9.setObjectName("horizontalLayout_9")
        self.label_2 = QtWidgets.QLabel(self.verticalLayoutWidget_2)
        self.label_2.setObjectName("label_2")
        self.horizontalLayout_9.addWidget(self.label_2)
        self.descriptionTextBox = QtWidgets.QLineEdit(self.verticalLayoutWidget_2)
        self.descriptionTextBox.setObjectName("descriptionTextBox")
        self.horizontalLayout_9.addWidget(self.descriptionTextBox)
        self.verticalLayout_3.addLayout(self.horizontalLayout_9)
        self.horizontalLayout_10 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_10.setSpacing(6)
        self.horizontalLayout_10.setObjectName("horizontalLayout_10")
        self.label_3 = QtWidgets.QLabel(self.verticalLayoutWidget_2)
        self.label_3.setObjectName("label_3")
        self.horizontalLayout_10.addWidget(self.label_3)
        self.eastTextBox = QtWidgets.QLineEdit(self.verticalLayoutWidget_2)
        self.eastTextBox.setObjectName("eastTextBox")
        self.horizontalLayout_10.addWidget(self.eastTextBox)
        self.verticalLayout_3.addLayout(self.horizontalLayout_10)
        self.horizontalLayout_8 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_8.setSpacing(6)
        self.horizontalLayout_8.setObjectName("horizontalLayout_8")
        self.label_4 = QtWidgets.QLabel(self.verticalLayoutWidget_2)
        self.label_4.setObjectName("label_4")
        self.horizontalLayout_8.addWidget(self.label_4)
        self.northTextBox = QtWidgets.QLineEdit(self.verticalLayoutWidget_2)
        self.northTextBox.setObjectName("northTextBox")
        self.horizontalLayout_8.addWidget(self.northTextBox)
        self.verticalLayout_3.addLayout(self.horizontalLayout_8)
        self.verticalLayoutWidget_3 = QtWidgets.QWidget(self.centralWidget)
        self.verticalLayoutWidget_3.setGeometry(QtCore.QRect(30, 280, 371, 61))
        self.verticalLayoutWidget_3.setObjectName("verticalLayoutWidget_3")
        self.verticalLayout_4 = QtWidgets.QVBoxLayout(self.verticalLayoutWidget_3)
        self.verticalLayout_4.setContentsMargins(11, 11, 11, 11)
        self.verticalLayout_4.setSpacing(6)
        self.verticalLayout_4.setObjectName("verticalLayout_4")
        self.label_6 = QtWidgets.QLabel(self.verticalLayoutWidget_3)
        self.label_6.setObjectName("label_6")
        self.verticalLayout_4.addWidget(self.label_6)
        self.horizontalLayout_11 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_11.setSpacing(6)
        self.horizontalLayout_11.setObjectName("horizontalLayout_11")
        self.lineEdit_2 = QtWidgets.QLineEdit(self.verticalLayoutWidget_3)
        self.lineEdit_2.setObjectName("lineEdit_2")
        self.horizontalLayout_11.addWidget(self.lineEdit_2)
        self.selectKMLButton = QtWidgets.QToolButton(self.verticalLayoutWidget_3)
        self.selectKMLButton.setObjectName("selectKMLButton")

        self.selectKMLButton.clicked.connect(self.kmlDialog)

        self.horizontalLayout_11.addWidget(self.selectKMLButton)
        self.verticalLayout_4.addLayout(self.horizontalLayout_11)
        self.exportKMLButton = QtWidgets.QPushButton(self.centralWidget)
        self.exportKMLButton.setGeometry(QtCore.QRect(290, 360, 113, 32))
        self.exportKMLButton.setObjectName("exportKMLButton")

        self.exportKMLButton.clicked.connect(self.export_click)
        self.exportKMLButton.setStyleSheet("QPushButton  {background: '#5D9CEC'; color:white; font-weight:bold;  font-family:'微軟正黑體';}");

        self.label_7 = QtWidgets.QLabel(self.centralWidget)
        self.label_7.setGeometry(QtCore.QRect(40, 360, 111, 16))
        self.label_7.setObjectName("label_7")
        MainWindow.setCentralWidget(self.centralWidget)
        self.menuBar = QtWidgets.QMenuBar(MainWindow)
        self.menuBar.setGeometry(QtCore.QRect(0, 0, 425, 22))
        self.menuBar.setObjectName("menuBar")
        MainWindow.setMenuBar(self.menuBar)
        self.mainToolBar = QtWidgets.QToolBar(MainWindow)
        self.mainToolBar.setObjectName("mainToolBar")
        MainWindow.addToolBar(QtCore.Qt.TopToolBarArea, self.mainToolBar)
        self.statusBar = QtWidgets.QStatusBar(MainWindow)
        self.statusBar.setObjectName("statusBar")
        self.statusBar.setStyleSheet("QStatusBar  {background: '#5D9CEC'; color:white; font-weight:bold;}");
        MainWindow.setStatusBar(self.statusBar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "ExcelToKML"))
        self.label_5.setText(_translate("MainWindow", "請選擇xlsx檔案："))
        self.selectExcelButton.setText(_translate("MainWindow", "..."))
        self.label.setText(_translate("MainWindow", "名稱欄位："))
        self.nameTextBox.setText(_translate("MainWindow", "站名"))
        self.label_2.setText(_translate("MainWindow", "描述欄位："))
        self.descriptionTextBox.setText(_translate("MainWindow", "所在鄉鎮"))
        self.label_3.setText(_translate("MainWindow", "經度欄位："))
        self.eastTextBox.setText(_translate("MainWindow", "E"))
        self.label_4.setText(_translate("MainWindow", "緯度欄位："))
        self.northTextBox.setText(_translate("MainWindow", "N"))
        self.label_6.setText(_translate("MainWindow", "請選擇 KML 匯出路徑："))
        self.selectKMLButton.setText(_translate("MainWindow", "..."))
        self.exportKMLButton.setText(_translate("MainWindow", "匯出 KML"))
        self.label_7.setText(_translate("MainWindow", ""))
        self.statusBar.showMessage('就緒')

    def export_click(self):
        self.statusBar.showMessage('匯出作業進行中...')

        name = self.nameTextBox.text()
        description = self.descriptionTextBox.text()
        east = self.eastTextBox.text()
        north = self.northTextBox.text()

        #print('{} {} {} {}'.format(name,description,east,north))
        self.statusBar.showMessage('匯出完成')

    def excelDialog(self):
        fname = QFileDialog.getOpenFileName(self, '選擇檔案')
        if fname[0]:
            self.lineEdit.setText(fname[0]) 

    def kmlDialog(self):
        fname = QFileDialog.getExistingDirectory(self, '選擇資料夾')
        if str(fname):
            self.lineEdit_2.setText(fname) 

class MainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super(MainWindow, self).__init__(parent)
        self.setupUi(self)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())