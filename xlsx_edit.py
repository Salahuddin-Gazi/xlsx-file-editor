from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QFileDialog, QWidget
import openpyxl, datetime, os

paths = []
class Ui_Dialog(QWidget):
    def setupUi(self, Dialog):
        Dialog.setObjectName("Dialog")
        Dialog.resize(533, 584)
        self.label = QtWidgets.QLabel(Dialog)
        self.label.setGeometry(QtCore.QRect(170, 0, 231, 21))
        font = QtGui.QFont()
        font.setFamily("Cambria")
        font.setPointSize(12)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.label_2 = QtWidgets.QLabel(Dialog)
        self.label_2.setGeometry(QtCore.QRect(10, 110, 261, 16))
        font = QtGui.QFont()
        font.setFamily("Cambria")
        font.setPointSize(10)
        self.label_2.setFont(font)
        self.label_2.setObjectName("label_2")

        # selectFiles_btn
        self.selectFiles_btn = QtWidgets.QPushButton(Dialog)
        self.selectFiles_btn.setGeometry(QtCore.QRect(10, 150, 151, 41))
        self.selectFiles_btn.setObjectName("selectFiles_btn")

        # showSelectedFiles - Box - 1
        self.showSelectedFiles = QtWidgets.QPlainTextEdit(Dialog)
        self.showSelectedFiles.setGeometry(QtCore.QRect(170, 150, 351, 161))
        self.showSelectedFiles.setObjectName("showSelectedFiles")

        # editFiles_btn
        self.editFiles_btn = QtWidgets.QPushButton(Dialog)
        self.editFiles_btn.setGeometry(QtCore.QRect(10, 330, 151, 41))
        self.editFiles_btn.setObjectName("editFiles_btn")

        # reset_btn
        self.reset_btn = QtWidgets.QPushButton(Dialog)
        self.reset_btn.setGeometry(QtCore.QRect(10, 500, 151, 41))
        self.reset_btn.setObjectName("reset_btn")

        # showEditedFiles - Box - 2
        self.showEditedFiles = QtWidgets.QPlainTextEdit(Dialog)
        self.showEditedFiles.setGeometry(QtCore.QRect(170, 330, 351, 161))
        self.showEditedFiles.setObjectName("showEditedFiles")

        self.retranslateUi(Dialog)
        QtCore.QMetaObject.connectSlotsByName(Dialog)

    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "Dialog"))
        self.label.setText(_translate("Dialog", "3G SiteAvailabilityReport Editor"))
        self.label_2.setText(_translate("Dialog", "Choose 3G SiteAvailabilityReport Files Only"))
        self.selectFiles_btn.setText(_translate("Dialog", "Select Files"))
        self.editFiles_btn.setText(_translate("Dialog", "Start Modifying"))
        self.reset_btn.setText(_translate("Dialog", "Reset"))

        # When selectFiles_btn button is clicked
        self.selectFiles_btn.clicked.connect(self.selectFiles)

        # When editFiles_btn button is clicked
        self.editFiles_btn.clicked.connect(self.editFiles)

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
            # self.showSelectedFiles.appendPlainText(file[-49:-5]+'\n')
            self.showSelectedFiles.appendPlainText(file+'\n')

    def editFiles(self):
        dir = os.getcwd() + "\\modified_files\\"
        modified_files = os.path.isdir(dir)
        if modified_files == False:
            os.mkdir(os.getcwd()+ "\\modified_files\\")
        if len(paths) > 0:
            print("Work in Progress ...")
            self.showEditedFiles.setPlainText("Work in Progress ...")
            for path in paths:
                # filename = path[-49:-5]
                date = path[-13:-5]
                year = int("20"+date[-2:])
                month = int(date[0:2])
                day = int(date[3:5]) - 2
                if day<0:
                    month = month - 1
                    if month in [1,3,5,7,8,10,12]:
                        day = 30
                    else:
                        day = 29
                elif day == 0:
                    month = month - 1
                    if month in [1,3,5,7,8,10,12]:
                        day = 31
                    else:
                        day = 30
                self.modifyMyExcelSheet(path, year, month, day)
            print("Finished ...")
            self.showEditedFiles.appendPlainText("Finished ...")

    def resetHandler (self):
        paths.clear()
        self.showEditedFiles.setPlainText("")
        self.showSelectedFiles.setPlainText("")


    def modifyMyExcelSheet(self, path, year, month, day):
        xlsx_file = path
        wb_obj = openpyxl.load_workbook(xlsx_file)
        sheet = wb_obj["RawData"]
        collectables = []
        
        for row in sheet.values:
            if 'Total Site Down' not in row:
                collectables.append(list(row))
        mayDayFst = datetime.datetime(year, month, day)
        mayDayLst = datetime.datetime(year, month, day, 23, 59, 59)

        for row in collectables:
            if type(row[9]) is type(mayDayLst):
                if (row[9] > mayDayLst or row[9] < mayDayFst):
                    row[9] = mayDayLst
            if type(row[7]) is type(mayDayFst):
                if row[7] < mayDayFst:
                    row[7] = mayDayFst

        myExport = openpyxl.Workbook()
        mySheet = myExport.active
        for row in collectables:
            myRow = tuple(row)
            mySheet.append(myRow)
        gg = "modified_file-3G_SiteAvailabilityReport_"+str(month)+"-"+str(day)+"-"+str(year)
        files_path= os.getcwd()+"\\modified_files\\"+gg+".xlsx"
        # files_path = QFileDialog.getSaveFileName(
        #     parent= self,
        #     caption= "Save Files", 
        #     directory= os.getcwd()+"/"+gg,
        #     filter= "XLSX/XLS Files (*.xlsx *.xls)"
        # )
        self.showEditedFiles.appendPlainText(gg)
        myExport.save(files_path)
        print(gg)

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Dialog = QtWidgets.QDialog()
    ui = Ui_Dialog()
    ui.setupUi(Dialog)
    Dialog.show()
    sys.exit(app.exec_())
