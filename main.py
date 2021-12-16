import sys

import pandas as pd
from PyQt5 import QtCore
from PyQt5 import QtWidgets
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QDialog, QApplication, QTableWidgetItem, QFileDialog, QMessageBox
from PyQt5.uic import loadUi


class App(QDialog):
    def __init__(self):
        super(App, self).__init__()
        loadUi("design5.ui", self)
        self.filename.setText("")
        self.pushButton_openfile.clicked.connect(self.browsefiles)
        self.pushButton_loadfile.clicked.connect(self.load_exclel_data)
        self.showMaximized()
        self.pushButton_script.clicked.connect(self.script)
        self.pushButton_exit.clicked.connect(QtCore.QCoreApplication.instance().quit)

    # load_exclel_file
    # In order to make pandas able to read .xlsx files, install openpyxl
    def load_exclel_data(self):
        excel_file_dir = self.filename.text()
        try:
            df = pd.read_excel(excel_file_dir)
            if df.size == 0:
                return

            # clear tableWidget
            self.tableWidget.setRowCount(0)
            self.tableWidget.setColumnCount(0)

            df.fillna('', inplace=True)
            self.tableWidget.setRowCount(df.shape[0])
            self.tableWidget.setColumnCount(df.shape[1])
            self.tableWidget.setHorizontalHeaderLabels(df.columns)

            # returns pandas array object
            for row in df.iterrows():
                values = row[1]
                for col_index, value in enumerate(values):
                    if isinstance(value, (float, int)):
                        value = '{0:0,.0f}'.format(value)
                    tabl_item = QTableWidgetItem(str(value))
                    self.tableWidget.setItem(row[0], col_index, tabl_item)
                    self.tableWidget.setRowHeight(row[0], 180)
            #self.tableWidget.resizeRowsToContents()
            self.tableWidget.setColumnWidth(0, 150)
            self.tableWidget.setColumnWidth(1, 200)
            self.tableWidget.setColumnWidth(2, 250)
            self.tableWidget.setColumnWidth(3, 400)
            self.tableWidget.setColumnWidth(4, 250)
            self.tableWidget.setColumnWidth(5, 250)

        except ValueError:
            QMessageBox.warning(self, "Error", "File Error", QMessageBox.Cancel)
            return None
        except FileNotFoundError:
            QMessageBox.warning(self, "Error", "File Not Found", QMessageBox.Cancel)
            return None

    #script_exclel file data into file\files .txt
    def script(self):
        try:
            # excel_file_dir = "data1.xlsx"
            excel_file_dir = self.filename.text()
            df = pd.read_excel(excel_file_dir)
            if df.size == 0:
                return

            # returns pandas array object
            for row in df.iterrows():
                letter_data = row[1]
                count_file_num = row[0] + 1
                greeting = letter_data[0]
                name = letter_data[1]
                company_name = letter_data[2]
                letter_body = letter_data[3]
                street_address = letter_data[4]
                city_address = letter_data[5]
                path = "output_data_files\\" + str(count_file_num) + ".txt"
                full_letter = greeting + " " + name + "," + " \n" + letter_body + "\n" + company_name + "\n" + street_address + "\n" + city_address
                file = open(path, 'w', encoding='utf-8')
                file.write(full_letter)
                file.close()
            # print(f"\n{count_file_num} {greeting} {name},\n{letter_body}\n{company_name}\n{street_address}\n{city_address}")

        except ValueError:
            QMessageBox.warning(self, "Error", "File Error", QMessageBox.Ok)
            return None
        except FileNotFoundError:
            QMessageBox.warning(self, "Error", "File Not Found", QMessageBox.Ok)
            return None

    # Browse files to select Excel file
    def browsefiles(self):
        fname = QFileDialog.getOpenFileName(self, 'Open file', 'C:', 'xlsx files, *.xlsx')
        self.filename.setText(fname[0])

app = QApplication(sys.argv)
mainwindow = App()
widget = QtWidgets.QStackedWidget()
widget.addWidget(mainwindow)
widget.setWindowFlag(Qt.WindowMinimizeButtonHint, True)
widget.setWindowFlag(Qt.WindowMaximizeButtonHint, True)
widget.setMinimumWidth(1000)
widget.setMinimumHeight(600)

widget.setWindowTitle("Excel viewer")
widget.show()
app.exec_()
