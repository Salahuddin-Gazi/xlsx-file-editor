from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QFileDialog, QWidget
import openpyxl, os

paths=[]
class Ui_Dialog(QWidget):
    def setupUi(self, Dialog):
        Dialog.setObjectName("Dialog")
        Dialog.resize(533, 620)
        self.label = QtWidgets.QLabel(Dialog)
        self.label.setGeometry(QtCore.QRect(200, 0, 161, 21))
        font = QtGui.QFont()
        font.setFamily("Cambria")
        font.setPointSize(12)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.label_2 = QtWidgets.QLabel(Dialog)
        self.label_2.setGeometry(QtCore.QRect(10, 80, 261, 16))
        font = QtGui.QFont()
        font.setFamily("Cambria")
        font.setPointSize(11)
        self.label_2.setFont(font)
        self.label_2.setObjectName("label_2")

        # selectFiles_btn
        self.selectFiles_btn = QtWidgets.QPushButton(Dialog)
        self.selectFiles_btn.setGeometry(QtCore.QRect(10, 120, 151, 41))
        self.selectFiles_btn.setObjectName("selectFiles_btn")
        
        # showSelectedFiles - Box - 1
        self.showSelectedFiles = QtWidgets.QPlainTextEdit(Dialog)
        self.showSelectedFiles.setGeometry(QtCore.QRect(170, 120, 351, 161))
        self.showSelectedFiles.setObjectName("showSelectedFiles")
        
        # mergeFiles_btn
        self.mergeFiles_btn = QtWidgets.QPushButton(Dialog)
        self.mergeFiles_btn.setGeometry(QtCore.QRect(10, 340, 151, 41))
        self.mergeFiles_btn.setObjectName("mergeFiles_btn")
        
        # mergedFileName - Box - 2
        self.mergedFileName = QtWidgets.QPlainTextEdit(Dialog)
        self.mergedFileName.setGeometry(QtCore.QRect(170, 340, 351, 41))
        self.mergedFileName.setObjectName("mergedFileName")
        
        # outputWindow - Box - 3
        self.outputWindow = QtWidgets.QPlainTextEdit(Dialog)
        self.outputWindow.setGeometry(QtCore.QRect(170, 410, 351, 161))
        self.outputWindow.setObjectName("outputWindow")

        # reset_btn
        self.reset_btn = QtWidgets.QPushButton(Dialog)
        self.reset_btn.setGeometry(QtCore.QRect(10, 575, 151, 41))
        self.reset_btn.setObjectName("reset_btn")

        self.label_3 = QtWidgets.QLabel(Dialog)
        self.label_3.setGeometry(QtCore.QRect(170, 310, 261, 21))
        font = QtGui.QFont()
        font.setFamily("Cambria")
        font.setPointSize(11)
        self.label_3.setFont(font)
        self.label_3.setObjectName("label_3")

        self.retranslateUi(Dialog)
        QtCore.QMetaObject.connectSlotsByName(Dialog)

    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "Dialog"))
        self.label.setText(_translate("Dialog", "Merge XLSX/XLS Files"))
        self.label_2.setText(_translate("Dialog", "Choose only xlsx/xls files"))
        self.selectFiles_btn.setText(_translate("Dialog", "Select Files"))
        self.mergeFiles_btn.setText(_translate("Dialog", "Merge Files"))
        self.outputWindow.setPlainText(_translate("Dialog", ">>>>>>Files Output Window<<<<<<"))
        self.label_3.setText(_translate("Dialog", "Please insert a name for merged files"))
        self.reset_btn.setText(_translate("Dialog", "Reset"))

        # When selectFiles_btn button is clicked
        self.selectFiles_btn.clicked.connect(self.selectFiles)

        # When mergeFiles_btn button is clicked
        self.mergeFiles_btn.clicked.connect(self.mergeFiles)

        # When reset_btn button is clicked
        self.reset_btn.clicked.connect(self.resetHandler)

    def selectFiles(self):
        filename = QFileDialog.getOpenFileNames(
            parent=self,
            caption="Select Files", 
            directory=os.getcwd(),
            filter="XLSX/XLS Files (*.xlsx *.xls)"
        )
        path = filename[0]
        self.showSelectedFiles.setPlainText("")
        for file in path:
            if file not in paths:
                paths.append(file)
            self.showSelectedFiles.appendPlainText(file+'\n')

    def mergeFiles(self):
        file_dir = os.getcwd() + "\\merged_file\\"
        merged_file = os.path.isdir(file_dir)
        if merged_file == False:
            os.mkdir(os.getcwd()+ "\\merged_file\\")
        fileName = self.mergedFileName.toPlainText()
        dir = paths
        if len(fileName) !=0 and len(dir) > 0:
            self.mergeExcelSheet(dir, fileName)

    def resetHandler (self):
        paths.clear()
        self.outputWindow.setPlainText(">>>>>>Files Output Window<<<<<<")
        self.showSelectedFiles.setPlainText("")
        self.mergedFileName.setPlainText("")

    def mergeExcelSheet(self, dir, fileName):
        print("Work in Progress ...")
        self.outputWindow.appendPlainText("Work in Progress ...")
        dest_wb = openpyxl.Workbook()
        collectables = []
        for file in dir:
            print(file)
            self.outputWindow.appendPlainText(file+"\n")
            xlsx_file = file
            wb_obj = openpyxl.load_workbook(xlsx_file)
            sheet = wb_obj.active

            for row in sheet.values:
                collectables.append(row)

        sheet = dest_wb.active
        for row in collectables:
            sheet.append(row)
        gg = fileName
        files_path= os.getcwd()+"\\merged_file\\"+gg+".xlsx"
        dest_wb.save(files_path)
        print("File Saved.... as"+' '+gg)
        self.outputWindow.appendPlainText("File Saved.... as"+' '+gg)

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Dialog = QtWidgets.QDialog()
    ui = Ui_Dialog()
    ui.setupUi(Dialog)
    Dialog.show()
    sys.exit(app.exec_())
