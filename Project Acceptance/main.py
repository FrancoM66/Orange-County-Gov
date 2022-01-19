from operator import concat
import os
from PyQt6 import QtCore, QtGui, QtWidgets
from PyQt6.QtWidgets import QMessageBox
import ctypes
import checkBox
import search
import excel
import createfolder

class Ui_MainWindow(object):
    
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("Excel Exterminator")
        MainWindow.setWindowModality(QtCore.Qt.WindowModality.WindowModal)
        MainWindow.setEnabled(True)
        self.user32 = ctypes.windll.user32  # custom
        MainWindow.resize(self.user32.GetSystemMetrics(
            0)-10, self.user32.GetSystemMetrics(1))  # custom
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.plan_file = QtWidgets.QFrame(self.centralwidget)
        self.plan_file.setGeometry(QtCore.QRect(10, 100, 153, 94))
        self.plan_file.setFrameShape(QtWidgets.QFrame.Shape.StyledPanel)
        self.plan_file.setFrameShadow(QtWidgets.QFrame.Shadow.Raised)
        self.plan_file.setObjectName("plan_file")
        self.planfile_label = QtWidgets.QLabel(self.plan_file)
        self.planfile_label.setGeometry(QtCore.QRect(20, 10, 54, 16))
        self.planfile_label.setObjectName("planfile_label")
        self.planfile_entry = QtWidgets.QLineEdit(self.plan_file)
        self.planfile_entry.setEnabled(False)
        self.planfile_entry.setGeometry(QtCore.QRect(20, 30, 113, 22))
        self.planfile_entry.setObjectName("planfile_entry")
        self.onlyInt = QtGui.QIntValidator()  # custom
        self.planfile_entry.setValidator(self.onlyInt)  # custom
        self.planfile_Btn = QtWidgets.QPushButton(self.plan_file)
        self.planfile_Btn.setEnabled(False)
        self.planfile_Btn.setGeometry(QtCore.QRect(30, 60, 91, 24))
        self.planfile_Btn.setAutoFillBackground(False)
        self.planfile_Btn.setObjectName("planfile_Btn")
        self.planfile_Btn.clicked.connect(lambda: search.search_clicked(self, self.planfile_entry.text(
        ), self.development_checkB, self.CIP_checkB, self.pump_folder, self.pressure_folder, self.gravity_folder))  # custom

        self.select_folder = QtWidgets.QFrame(self.centralwidget)
        self.select_folder.setGeometry(QtCore.QRect(10, 290, 171, 181))
        self.select_folder.setFrameShape(QtWidgets.QFrame.Shape.StyledPanel)
        self.select_folder.setFrameShadow(QtWidgets.QFrame.Shadow.Raised)
        self.select_folder.setObjectName("select_folder")
        self.open_folder_label = QtWidgets.QLabel(self.select_folder)
        self.open_folder_label.setGeometry(QtCore.QRect(20, 10, 84, 16))
        self.open_folder_label.setObjectName("open_folder_label")
        self.pump_checkB = QtWidgets.QCheckBox(self.select_folder)
        self.pump_checkB.setEnabled(False)
        self.pump_checkB.setGeometry(QtCore.QRect(30, 40, 75, 20))
        self.pump_checkB.setObjectName("pump_checkB")
        self.pressure_checkB = QtWidgets.QCheckBox(self.select_folder)
        self.pressure_checkB.setEnabled(False)
        self.pressure_checkB.setGeometry(QtCore.QRect(30, 60, 75, 20))
        self.pressure_checkB.setObjectName("pressure_checkB")
        self.gravity_checkB = QtWidgets.QCheckBox(self.select_folder)
        self.gravity_checkB.setEnabled(False)
        self.gravity_checkB.setGeometry(QtCore.QRect(30, 80, 75, 20))
        self.gravity_checkB.setObjectName("gravity_checkB")
        self.options_pane = QtWidgets.QFrame(self.centralwidget)
        self.options_pane.setGeometry(QtCore.QRect(10, 520, 141, 261))
        self.options_pane.setFrameShape(QtWidgets.QFrame.Shape.StyledPanel)
        self.options_pane.setFrameShadow(QtWidgets.QFrame.Shadow.Raised)
        self.options_pane.setObjectName("options_pane")
        self.options_Label = QtWidgets.QLabel(self.options_pane)
        self.options_Label.setGeometry(QtCore.QRect(20, 10, 42, 16))
        self.options_Label.setObjectName("options_Label")
        self.open_existing_Btn = QtWidgets.QPushButton(self.options_pane)
        self.open_existing_Btn.setEnabled(False)
        self.open_existing_Btn.setGeometry(QtCore.QRect(30, 40, 91, 24))
        self.open_existing_Btn.setObjectName("open_existing_Btn")

        self.open_existing_Btn.clicked.connect(lambda: excel.init_table(
            self, self.tableview, self.concat, self.pump_checkB, self.gravity_checkB, self.pressure_checkB, self.CIP_checkB, self.development_checkB, "",""))

        
        # self.open_existing_Btn.clicked.connect(lambda: test2.loadExcelData(self, self.tableview, self.concat, self.pump_checkB, self.gravity_checkB, self.pressure_checkB, self.CIP_checkB, self.development_checkB))
        
        self.add_row = QtWidgets.QPushButton(self.options_pane)
        self.add_row.setEnabled(False)
        self.add_row.setGeometry(QtCore.QRect(30, 160, 91, 24))
        self.add_row.setObjectName("add_row")
        self.add_row.clicked.connect(lambda: excel.add_rows(self,self.tableview))
        self.create_new_Btn = QtWidgets.QPushButton(self.options_pane)
        self.create_new_Btn.setEnabled(False)
        self.create_new_Btn.setGeometry(QtCore.QRect(30, 70, 91, 24))
        self.create_new_Btn.setObjectName("create_new_Btn")
        self.create_new_Btn.clicked.connect(lambda: createfolder.create_new(self, self.concat))
        self.create_pdf = QtWidgets.QPushButton(self.options_pane)
        self.create_pdf.setEnabled(False)
        self.create_pdf.setGeometry(QtCore.QRect(30, 100, 91, 24))
        self.create_pdf.setObjectName("create_pdf")
        self.save_btn = QtWidgets.QPushButton(self.options_pane)
        self.save_btn.setEnabled(False)
        self.save_btn.setGeometry(QtCore.QRect(30, 130, 91, 24))
        self.save_btn.setObjectName("save_btn")
        self.save_btn.clicked.connect(lambda: excel.check_b4_save(self))
        self.reset_btn = QtWidgets.QPushButton(self.options_pane)
        self.reset_btn.setEnabled(True)
        self.reset_btn.setGeometry(QtCore.QRect(30, 218, 91, 24))
        self.reset_btn.setObjectName("reset_btn")
        self.reset_btn.clicked.connect(lambda: self.reset_all())
        self.work_area = QtWidgets.QFrame(self.centralwidget)
        self.work_area.setGeometry(QtCore.QRect(10, 10, 153, 94))
        self.work_area.setFrameShape(QtWidgets.QFrame.Shape.StyledPanel)
        self.work_area.setFrameShadow(QtWidgets.QFrame.Shadow.Raised)
        self.work_area.setObjectName("work_area")
        self.work_entry = QtWidgets.QLineEdit(self.work_area)
        self.work_entry.setGeometry(QtCore.QRect(20, 30, 113, 22))
        self.work_entry.setObjectName("work_entry")
        self.label_2 = QtWidgets.QLabel(self.work_area)
        self.label_2.setGeometry(QtCore.QRect(20, 10, 55, 16))
        self.label_2.setObjectName("label_2")
        self.work_entry_Btn = QtWidgets.QPushButton(self.work_area)
        self.work_entry_Btn.setGeometry(QtCore.QRect(30, 60, 91, 24))
        self.work_entry_Btn.setAutoFillBackground(False)
        self.work_entry_Btn.setAutoDefault(False)
        self.work_entry_Btn.setObjectName("work_entry_Btn")
        self.work_entry_Btn.clicked.connect(lambda: checkBox.check_work_area(
            self, self.pump_checkB, self.pressure_checkB, self.gravity_checkB, self.planfile_Btn, self.planfile_entry, self.work_entry.text()))  # custom
        self.DorP_frame = QtWidgets.QFrame(self.centralwidget)
        self.DorP_frame.setGeometry(QtCore.QRect(10, 190, 146, 101))
        self.DorP_frame.setFrameShape(QtWidgets.QFrame.Shape.StyledPanel)
        self.DorP_frame.setFrameShadow(QtWidgets.QFrame.Shadow.Raised)
        self.DorP_frame.setObjectName("DorP_frame")
        self.DorP_label = QtWidgets.QLabel(self.DorP_frame)
        self.DorP_label.setGeometry(QtCore.QRect(20, 10, 106, 16))
        self.DorP_label.setObjectName("DorP_label")
        self.development_checkB = QtWidgets.QCheckBox(self.DorP_frame)
        self.development_checkB.setEnabled(False)
        self.development_checkB.setGeometry(QtCore.QRect(30, 40, 94, 20))
        self.development_checkB.setObjectName("development_checkB")
        self.CIP_checkB = QtWidgets.QCheckBox(self.DorP_frame)
        self.CIP_checkB.setEnabled(False)
        self.CIP_checkB.setGeometry(QtCore.QRect(30, 60, 41, 20))
        self.CIP_checkB.setObjectName("CIP_checkB")
        self.enabled = self.CIP_checkB.setEnabled(False)
        self.development_checkB.stateChanged.connect(
            self.onStateChange)  # custom
        self.CIP_checkB.stateChanged.connect(self.onStateChange)  # custom

        self.tableview = QtWidgets.QTableView(self.centralwidget)
        self.tableview.setGeometry(QtCore.QRect(190, 10, self.user32.GetSystemMetrics(
            0) - 225, self.user32.GetSystemMetrics(1) - 300))  # custom
        self.tableview.setObjectName("tableView")

        

        self.select_folder_3 = QtWidgets.QFrame(self.centralwidget)
        self.select_folder_3.setGeometry(QtCore.QRect(10, 410, 125, 110))
        self.select_folder_3.setFrameShape(QtWidgets.QFrame.Shape.StyledPanel)
        self.select_folder_3.setFrameShadow(QtWidgets.QFrame.Shadow.Raised)
        self.select_folder_3.setObjectName("select_folder_3")
        self.open_folder_label_3 = QtWidgets.QLabel(self.select_folder_3)
        self.open_folder_label_3.setGeometry(QtCore.QRect(20, 10, 84, 16))
        self.open_folder_label_3.setObjectName("open_folder_label_3")
        self.pump_folder = QtWidgets.QCheckBox(self.select_folder_3)
        self.pump_folder.setEnabled(False)
        self.pump_folder.setGeometry(QtCore.QRect(30, 40, 75, 20))
        self.pump_folder.setObjectName("pump_folder")
        self.pressure_folder = QtWidgets.QCheckBox(self.select_folder_3)
        self.pressure_folder.setEnabled(False)
        self.pressure_folder.setGeometry(QtCore.QRect(30, 60, 75, 20))
        self.pressure_folder.setObjectName("pressure_folder")
        self.gravity_folder = QtWidgets.QCheckBox(self.select_folder_3)
        self.gravity_folder.setEnabled(False)
        self.gravity_folder.setGeometry(QtCore.QRect(30, 80, 75, 20))
        self.gravity_folder.setObjectName("gravity_folder")
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 1400, 22))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        self.open_existing_Btn.clicked.connect(lambda:self.reset_btn.setEnabled(False))
        self.open_existing_Btn.clicked.connect(lambda:self.work_entry_Btn.setEnabled(False))
        self.open_existing_Btn.clicked.connect(lambda:self.planfile_Btn.setEnabled(False))
        self.open_existing_Btn.clicked.connect(lambda:self.add_row.setEnabled(True))
        self.open_existing_Btn.clicked.connect(lambda:self.create_pdf.setEnabled(True))
        self.open_existing_Btn.clicked.connect(lambda:self.save_btn.setEnabled(True))
        
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)
        MainWindow.setTabOrder(self.work_entry, self.work_entry_Btn)
        MainWindow.setTabOrder(self.work_entry_Btn, self.planfile_entry)
        MainWindow.setTabOrder(self.planfile_entry, self.planfile_Btn)
        MainWindow.setTabOrder(self.planfile_Btn, self.pump_checkB)
        MainWindow.setTabOrder(self.pump_checkB, self.pressure_checkB)
        MainWindow.setTabOrder(self.pressure_checkB, self.gravity_checkB)
        MainWindow.setTabOrder(self.gravity_checkB, self.open_existing_Btn)
        MainWindow.setTabOrder(self.open_existing_Btn, self.create_new_Btn)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.planfile_label.setText(_translate("MainWindow", "Plan File #"))
        self.planfile_Btn.setText(_translate("MainWindow", "Enter"))
        self.open_folder_label.setText(
            _translate("MainWindow", "Active Category"))
        self.pump_checkB.setText(_translate("MainWindow", "Pump"))
        self.pressure_checkB.setText(_translate("MainWindow", "Pressure"))
        self.gravity_checkB.setText(_translate("MainWindow", "Gravity"))
        self.options_Label.setText(_translate("MainWindow", "Options"))
        self.open_existing_Btn.setText(
            _translate("MainWindow", "Open Existing"))
        self.add_row.setText(
            _translate("MainWindow", "Add Row"))
        self.create_new_Btn.setText(_translate("MainWindow", "Create New"))
        self.create_pdf.setText(_translate("MainWindow", "Create PDF"))
        self.save_btn.setText(_translate("MainWindow", "Save Changes"))
        self.reset_btn.setText(_translate("MainWindow", "Reset All Fields"))
        self.label_2.setText(_translate("MainWindow", "Work Area"))
        self.work_entry_Btn.setText(_translate("MainWindow", "Enter"))
        self.DorP_label.setText(_translate("MainWindow", "Development or CIP"))
        self.development_checkB.setText(
            _translate("MainWindow", "Development"))
        self.CIP_checkB.setText(_translate("MainWindow", "CIP"))
        self.open_folder_label_3.setText(
            _translate("MainWindow", "Folders Found"))
        self.pump_folder.setText(_translate("MainWindow", "Pump"))
        self.pressure_folder.setText(_translate("MainWindow", "Pressure"))
        self.gravity_folder.setText(_translate("MainWindow", "Gravity"))

    # custom
    def onStateChange(self):
        if self.CIP_checkB.isChecked():
            self.development_checkB.setEnabled(False)
            self.development_checkB.setChecked(False)
            self.CIP_checkB.setEnabled(False)

        if self.development_checkB.isChecked():
            self.CIP_checkB.setEnabled(False)
            self.CIP_checkB.setChecked(False)
            self.development_checkB.setEnabled(False)

        if self.pump_checkB.isChecked() and self.pump_folder.isChecked():

            self.open_existing_Btn.setEnabled(True)
            self.create_new_Btn.setEnabled(False)
            

        elif self.gravity_checkB.isChecked() and self.gravity_folder.isChecked():
            self.open_existing_Btn.setEnabled(True)
            self.create_new_Btn.setEnabled(False)
            

        elif self.pressure_checkB.isChecked() and self.pressure_folder.isChecked():

            self.open_existing_Btn.setEnabled(True)
            self.create_new_Btn.setEnabled(False)
            

        else:
            self.open_existing_Btn.setEnabled(False)
            self.create_new_Btn.setEnabled(True)
            self.create_pdf.setEnabled(False)
            self.save_btn.setEnabled(False)
            self.add_row.setEnabled(False)

    
    def reset_all(self):
        self.development_checkB.setChecked(False)
        self.development_checkB.setEnabled(False)
        self.CIP_checkB.setChecked(False)
        self.CIP_checkB.setEnabled(False)
        self.work_entry.clear()
        self.planfile_entry.clear()
        self.save_btn.setEnabled(False)
        self.planfile_entry.setEnabled(False)
        self.create_pdf.setEnabled(False)
        self.create_new_Btn.setEnabled(False)
        self.pressure_checkB.setChecked(False)
        self.gravity_checkB.setChecked(False)
        self.pump_checkB.setChecked(False)
        self.pressure_checkB.setEnabled(False)
        self.gravity_checkB.setEnabled(False)
        self.pump_checkB.setEnabled(False)
        self.pump_folder.setChecked(False)
        self.gravity_folder.setChecked(False)
        self.pressure_folder.setChecked(False)
        

    
        
def main():
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec())

            
if __name__ == "__main__":
    main()