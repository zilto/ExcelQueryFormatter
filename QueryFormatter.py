from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox, QWidget, QPushButton, QLineEdit, QFileDialog

try:
    myappid = "ThierryJean_QueryFormatter_1.0"
    PyQt5.QtWinExtras.QtWin.setCurrentProcessExplicitAppUserModelID(myappid)
    QtWin.setCurrentProcessExplicitAppUserModelID(myappid)
except:
    pass

import sys
from os import path
import re
import pandas as pd
import icon_resource

#finds the absolute path of the directory of the script
CUR_DIR = path.dirname(path.realpath(__file__))

class MyWindow(QMainWindow):
    def __init__(self):
        #initiates instance of MainWindow
        QMainWindow.__init__(self)
        #calls setupUi function to give a ui to MainWindow object
        self.setupUi(self)

    def setupUi(self, QueryFormatter):
        #creates the main window and gives it a name
        QueryFormatter.setObjectName("QueryFormatter")
        #sets fixed/not resizable by user MainWindow size
        QueryFormatter.setFixedSize(680, 180)
        #theme/design of the MainWindow
        #starts the layout of the components within a mainFrame
        QueryFormatter.setTabShape(QtWidgets.QTabWidget.Rounded)
        self.centralwidget = QtWidgets.QWidget(QueryFormatter)
        self.centralwidget.setObjectName("centralwidget")
        self.gridLayout = QtWidgets.QGridLayout(self.centralwidget)
        self.gridLayout.setObjectName("gridLayout")
        self.mainFrame = QtWidgets.QFrame(self.centralwidget)
        self.mainFrame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.mainFrame.setFrameShadow(QtWidgets.QFrame.Raised)
        self.mainFrame.setObjectName("mainFrame")
        self.mainSplitter = QtWidgets.QSplitter(self.mainFrame)
        self.mainSplitter.setGeometry(QtCore.QRect(20, 10, 610, 114))
        self.mainSplitter.setOrientation(QtCore.Qt.Vertical)
        self.mainSplitter.setObjectName("mainSplitter")
        self.layoutWidget = QtWidgets.QWidget(self.mainSplitter)
        self.layoutWidget.setObjectName("layoutWidget")
        self.horizontalLayout = QtWidgets.QHBoxLayout(self.layoutWidget)
        self.horizontalLayout.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.ogFname_label = QtWidgets.QLabel(self.layoutWidget)
        self.ogFname_label.setEnabled(True)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        #makes the size of the line edits fixed
        sizePolicy.setHeightForWidth(self.ogFname_label.sizePolicy().hasHeightForWidth())
        self.ogFname_label.setSizePolicy(sizePolicy)
        self.ogFname_label.setMinimumSize(QtCore.QSize(105, 0))
        self.ogFname_label.setObjectName("ogFname_label")
        self.horizontalLayout.addWidget(self.ogFname_label)
        self.ogFname_ledit = QtWidgets.QLineEdit(self.layoutWidget)
        self.ogFname_ledit.setObjectName("ogFname_ledit")
        self.horizontalLayout.addWidget(self.ogFname_ledit)
        self.ogFname_bBrowse = QtWidgets.QPushButton(self.layoutWidget)
        self.ogFname_bBrowse.setObjectName("ogFname_bBrowse")
        self.horizontalLayout.addWidget(self.ogFname_bBrowse)
        self.layoutWidget1 = QtWidgets.QWidget(self.mainSplitter)
        self.layoutWidget1.setObjectName("layoutWidget1")
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout(self.layoutWidget1)
        self.horizontalLayout_2.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.modelFname_label = QtWidgets.QLabel(self.layoutWidget1)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        #makes the size of the line edits fixed
        sizePolicy.setHeightForWidth(self.modelFname_label.sizePolicy().hasHeightForWidth())
        self.modelFname_label.setSizePolicy(sizePolicy)
        self.modelFname_label.setMinimumSize(QtCore.QSize(105, 0))
        self.modelFname_label.setObjectName("modelFname_label")
        self.horizontalLayout_2.addWidget(self.modelFname_label)
        self.modelFname_ledit = QtWidgets.QLineEdit(self.layoutWidget1)
        self.modelFname_ledit.setObjectName("modelFname_ledit")
        self.horizontalLayout_2.addWidget(self.modelFname_ledit)
        self.modelFname_bBrowse = QtWidgets.QPushButton(self.layoutWidget1)
        self.modelFname_bBrowse.setObjectName("modelFname_bBrowse")
        self.horizontalLayout_2.addWidget(self.modelFname_bBrowse)
        self.layoutWidget2 = QtWidgets.QWidget(self.mainSplitter)
        self.layoutWidget2.setObjectName("layoutWidget2")
        self.horizontalLayout_3 = QtWidgets.QHBoxLayout(self.layoutWidget2)
        self.horizontalLayout_3.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_3.setObjectName("horizontalLayout_3")
        self.outputDir_label = QtWidgets.QLabel(self.layoutWidget2)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        #makes the size of the line edits fixed
        sizePolicy.setHeightForWidth(self.outputDir_label.sizePolicy().hasHeightForWidth())
        self.outputDir_label.setSizePolicy(sizePolicy)
        self.outputDir_label.setMinimumSize(QtCore.QSize(105, 0))
        self.outputDir_label.setObjectName("outputDir_label")
        self.horizontalLayout_3.addWidget(self.outputDir_label)
        self.outputDir_ledit = QtWidgets.QLineEdit(self.layoutWidget2)
        self.outputDir_ledit.setObjectName("outputDir_ledit")
        self.horizontalLayout_3.addWidget(self.outputDir_ledit)
        self.outputDir_bBrowse = QtWidgets.QPushButton(self.layoutWidget2)
        self.outputDir_bBrowse.setObjectName("outputDir_bBrowse")
        self.horizontalLayout_3.addWidget(self.outputDir_bBrowse)
        self.layoutWidget3 = QtWidgets.QWidget(self.mainSplitter)
        self.layoutWidget3.setObjectName("layoutWidget3")
        self.horizontalLayout_4 = QtWidgets.QHBoxLayout(self.layoutWidget3)
        self.horizontalLayout_4.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_4.setObjectName("horizontalLayout_4")
        self.outputFname_Label = QtWidgets.QLabel(self.layoutWidget3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.outputFname_Label.sizePolicy().hasHeightForWidth())
        self.outputFname_Label.setSizePolicy(sizePolicy)
        self.outputFname_Label.setMinimumSize(QtCore.QSize(105, 0))
        self.outputFname_Label.setObjectName("outputFname_Label")
        self.horizontalLayout_4.addWidget(self.outputFname_Label)
        self.outputFname_ledit = QtWidgets.QLineEdit(self.layoutWidget3)
        self.outputFname_ledit.setObjectName("outputFname_ledit")
        self.horizontalLayout_4.addWidget(self.outputFname_ledit)
        self.formatButton = QtWidgets.QPushButton(self.layoutWidget3)
        self.formatButton.setObjectName("formatButton")
        self.horizontalLayout_4.addWidget(self.formatButton)
        self.gridLayout.addWidget(self.mainFrame, 0, 0, 1, 1)
        QueryFormatter.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(QueryFormatter)
        self.statusbar.setObjectName("statusbar")
        QueryFormatter.setStatusBar(self.statusbar)

        #connects the buttons to their function
        self.retranslateUi(QueryFormatter)
        self.ogFname_bBrowse.clicked.connect(lambda: self.clickBrowse(self.ogFname_ledit))
        self.modelFname_bBrowse.clicked.connect(lambda: self.clickBrowse(self.modelFname_ledit))
        self.outputDir_bBrowse.clicked.connect(lambda: self.clickBrowseDir(self.outputDir_ledit))
        self.formatButton.clicked.connect(self.clickFormat)

    def retranslateUi(self, QueryFormatter):
        _translate = QtCore.QCoreApplication.translate
        QueryFormatter.setWindowTitle(_translate("QueryFormatter", "MainWindow"))
        self.ogFname_label.setText(_translate("QueryFormatter", "Original File Name"))
        self.ogFname_bBrowse.setText(_translate("QueryFormatter", "Browse"))
        self.modelFname_label.setText(_translate("QueryFormatter", "Model File Name"))
        self.modelFname_bBrowse.setText(_translate("QueryFormatter", "Browse"))
        self.outputDir_label.setText(_translate("QueryFormatter", "Output Directory"))
        self.outputDir_bBrowse.setText(_translate("QueryFormatter", "Browse"))
        self.outputFname_Label.setText(_translate("QueryFormatter", "Output File Name"))
        self.formatButton.setText(_translate("QueryFormatter", "Format"))

    #file browsing method for buttons; writes file directory to line edit
    def clickBrowse(self, lineEdit):
        dialog = QFileDialog()
        fname, filter = dialog.getOpenFileName(parent=None, caption="Select file", filter="Excel files (*.xlsx)")
        lineEdit.setText(fname)

    #directory browsing method for buttons; writes directory to line edit
    def clickBrowseDir(self, lineEdit):
        dialog = QFileDialog()
        dirname = dialog.getExistingDirectory(self, "Select a directory")
        lineEdit.setText(dirname)

    #formatting trigger for button; checks validate_request first
    def clickFormat(self):
        #validates request
        if self.validate_request():
            #triggers query formatting function
            format_query(self.ogFname_ledit.text(), self.modelFname_ledit.text(), self.outputDir_ledit.text(), self.outputFname_ledit.text())
            #popup to indicate the formatting is completed
            self.popupFormat()
            #clears the output file name lineedit
            self.outputFname_ledit.clear()

    #validation before formatting; raises a popup if request incorrect/incomplete
    def validate_request(self):
        #generic messages to be formatted
        missing = "{} is missing.\nPlease provided one."
        cant_find = "{} cannot be found.\nPlease indicate another one."
        already_exists = "{} already exists in this directory.\nPlease indicate another one."
        #original file name field is empty
        if not self.ogFname_ledit.text():
            self.popupRequestError(missing.format("Original file"))
        #model file name field is empty
        elif not self.modelFname_ledit.text():
            self.popupRequestError(missing.format("Model file"))
        #output file directory field is empty
        elif not self.outputDir_ledit.text():
            self.popupRequestError(missing.format("Output directory"))
        #output file name field is empty
        elif not self.outputFname_ledit.text():
            self.popupRequestError(missing.format("Output file name"))

        #original file doesn't exist a location indicated
        elif not path.isfile(self.ogFname_ledit.text()):
            self.popupRequestError(cant_find.format("Original file"))
        #model file doesn't exist at location indicated
        elif not path.isfile(self.modelFname_ledit.text()):
            self.popupRequestError(cant_find.format("Model file"))
        #output directory doesn't exist at location indicated
        elif not path.isdir(self.outputDir_ledit.text()):
            self.popupRequestError(cant_find.format("Output directory"))
        #file with output file name already exists in output directory
        elif path.isfile(self.outputDir_ledit.text() + "/" + self.outputFname_ledit.text() + ".xlsx"):
            self.popupRequestError(already_exists.format(self.outputFname_ledit.text()))
        #returns True if the request is valid (no flag raised)
        else:
            return True

    def popupFormat(self):
        msg = QMessageBox()
        msg.setWindowTitle("File formatted")
        msg.setText("Your file has been formatted successfully.")
        msg.setIcon(QMessageBox.Information)
        msg.setStandardButtons(QMessageBox.Ok)
        x = msg.exec_()

    def popupRequestError(self, error_msg):
        msg = QMessageBox()
        msg.setWindowTitle("Request Error")
        msg.setText(error_msg)
        msg.setIcon(QMessageBox.Critical)
        msg.setStandardButtons(QMessageBox.Ok)
        x = msg.exec_()

def format_query(excel1, excel2, outputdir, outputname):
    p = re.compile("^(?!(Unnamed))")
    og_query = pd.read_excel(excel1, header=0)
    model_query = pd.read_excel(excel2, header=0)
    model_cols = model_query.columns.tolist()
    filtered_cols = [col for col in model_cols if p.match(col)]
    formatted_query = og_query.filter(filtered_cols)
    formatted_query.to_excel(outputdir + "/" + outputname + ".xlsx", index=False)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setWindowIcon(QtGui.QIcon(":/icons/icon_green.ico"))
    window = MyWindow()
    window.show()
    sys.exit(app.exec_())
