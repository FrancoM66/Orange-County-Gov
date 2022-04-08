#!/usr/bin/env python3

from PyQt6 import QtCore, QtGui, QtWidgets
import ctypes
import checkBox as checkBox
from search import *
from excel_create import *
from createfolder import *
from checkBox import *

class UI_Main_Window(object):

    def __init__(self, MainWindow):
        super().__init__()
        # checking screen size
        user32 = ctypes.windll.user32
        screen_size_x = user32.GetSystemMetrics(0)
        screen_size_y = user32.GetSystemMetrics(1)

        # setting up validator
        self.onlyInt = QtGui.QIntValidator() 

        # naming MainWindow object
        MainWindow.setObjectName("Excel Exterminator")
        MainWindow.setWindowModality(QtCore.Qt.WindowModality.WindowModal)
        MainWindow.setEnabled(True)
        MainWindow.setWindowIcon(QtGui.QIcon('O:\Field Services Division\Field Support Center\Project Acceptance\PA Excel Exterminator\imgs/Logo.jpg'))
        MainWindow.resize(int(screen_size_x) - 10, int(screen_size_y))

        # creating central widget
        self.central_widget =  QtWidgets.QWidget(MainWindow)
        self.central_widget.setObjectName("Central Widget")
        MainWindow.setCentralWidget(self.central_widget)

        # work area 
        self.work_area = QtWidgets.QFrame(self.central_widget)
        self.work_area.setGeometry(QtCore.QRect(10, 10, 153, 94))
        self.work_area.setFrameShadow(QtWidgets.QFrame.Shadow.Raised)
        self.work_area.setObjectName("work_area")

        # work area entry
        self.work_entry = QtWidgets.QLineEdit(self.work_area)
        self.work_entry.setGeometry(QtCore.QRect(20, 30, 113, 22))
        self.work_entry.setObjectName("work_entry")

        # work area label
        self.work_area_label = QtWidgets.QLabel(self.work_area)
        self.work_area_label.setGeometry(QtCore.QRect(20, 10, 55, 16))
        self.work_area_label.setObjectName("work_area_label")

        # work area button
        self.work_entry_Btn = QtWidgets.QPushButton(self.work_area)
        self.work_entry_Btn.setGeometry(QtCore.QRect(30, 60, 91, 24))
        self.work_entry_Btn.setAutoFillBackground(False)
        self.work_entry_Btn.setAutoDefault(False)
        self.work_entry_Btn.setObjectName("work_entry_Btn")

        # creating planfile on screen
        self.plan_file = QtWidgets.QFrame(self.central_widget)
        self.plan_file.setGeometry(QtCore.QRect(10, 100, 153, 90))
        self.plan_file.setFrameShadow(QtWidgets.QFrame.Shadow.Raised)
        self.plan_file.setObjectName("plan_file")

        # setting label inside the planfile
        self.planfile_label = QtWidgets.QLabel(self.plan_file)
        self.planfile_label.setGeometry(QtCore.QRect(20, 10, 54, 16))
        self.planfile_label.setObjectName("planfile_label")

        # setting the entry field into planfile
        self.planfile_entry = QtWidgets.QLineEdit(self.plan_file)
        self.planfile_entry.setEnabled(False)
        self.planfile_entry.setGeometry(QtCore.QRect(20, 30, 113, 22))
        self.planfile_entry.setObjectName("planfile_entry")
        self.planfile_entry.setValidator(self.onlyInt)

        # setting button into planfile
        self.planfile_Btn = QtWidgets.QPushButton(self.plan_file)
        self.planfile_Btn.setEnabled(False)
        self.planfile_Btn.setGeometry(QtCore.QRect(30, 60, 91, 24))
        self.planfile_Btn.setAutoFillBackground(False)
        self.planfile_Btn.setObjectName("planfile_Btn")

        # creating selected folder on screen
        self.select_folder = QtWidgets.QFrame(self.central_widget)
        self.select_folder.setGeometry(QtCore.QRect(10, 290, 171, 181))
        self.select_folder.setFrameShadow(QtWidgets.QFrame.Shadow.Raised)
        self.select_folder.setObjectName("select_folder")

        # label for open folder
        self.open_folder_label = QtWidgets.QLabel(self.select_folder)
        self.open_folder_label.setGeometry(QtCore.QRect(20, 10, 84, 16))
        self.open_folder_label.setObjectName("open_folder_label")

        # Check boxes into the selected folder 
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

        # options frame set on screen
        self.options_pane = QtWidgets.QFrame(self.central_widget)
        self.options_pane.setGeometry(QtCore.QRect(10, 520, 141, 300))
        self.options_pane.setFrameShadow(QtWidgets.QFrame.Shadow.Raised)
        self.options_pane.setObjectName("options_pane")

        # options label
        self.options_Label = QtWidgets.QLabel(self.options_pane)
        self.options_Label.setGeometry(QtCore.QRect(20, 10, 42, 16))
        self.options_Label.setObjectName("options_Label")

        # Development of CIP frame
        self.Dev_or_CIP_frame = QtWidgets.QFrame(self.central_widget)
        self.Dev_or_CIP_frame.setGeometry(QtCore.QRect(10, 190, 146, 101))
        self.Dev_or_CIP_frame.setFrameShadow(QtWidgets.QFrame.Shadow.Raised)
        self.Dev_or_CIP_frame.setObjectName("Dev_or_CIP_frame")

        # Development or CIP label
        self.Dev_or_CIP_label = QtWidgets.QLabel(self.Dev_or_CIP_frame)
        self.Dev_or_CIP_label.setGeometry(QtCore.QRect(20, 10, 106, 16))
        self.Dev_or_CIP_label.setObjectName("DorP_label")

        # Development Check Box
        self.development_checkB = QtWidgets.QCheckBox(self.Dev_or_CIP_frame)
        self.development_checkB.setEnabled(False)
        self.development_checkB.setGeometry(QtCore.QRect(30, 40, 94, 20))
        self.development_checkB.setObjectName("development_checkB")

        # CIP checkbox
        self.CIP_checkB = QtWidgets.QCheckBox(self.Dev_or_CIP_frame)
        self.CIP_checkB.setEnabled(False)
        self.CIP_checkB.setGeometry(QtCore.QRect(30, 60, 41, 20))
        self.CIP_checkB.setObjectName("CIP_checkB")

        # Select folder frame
        self.select_folder = QtWidgets.QFrame(self.central_widget)
        self.select_folder.setGeometry(QtCore.QRect(10, 410, 125, 110))
        self.select_folder.setFrameShadow(QtWidgets.QFrame.Shadow.Raised)
        self.select_folder.setObjectName("select_folder")

        # open folder label
        self.active_folder_label = QtWidgets.QLabel(self.select_folder)
        self.active_folder_label.setGeometry(QtCore.QRect(20, 10, 84, 16))
        self.active_folder_label.setObjectName("active_folder_label")

        # pump folder
        self.pump_folder = QtWidgets.QCheckBox(self.select_folder)
        self.pump_folder.setEnabled(False)
        self.pump_folder.setGeometry(QtCore.QRect(30, 40, 75, 20))
        self.pump_folder.setObjectName("pump_folder")

        # pressure folder
        self.pressure_folder = QtWidgets.QCheckBox(self.select_folder)
        self.pressure_folder.setEnabled(False)
        self.pressure_folder.setGeometry(QtCore.QRect(30, 60, 75, 20))
        self.pressure_folder.setObjectName("pressure_folder")

        # gravity folder
        self.gravity_folder = QtWidgets.QCheckBox(self.select_folder)
        self.gravity_folder.setEnabled(False)
        self.gravity_folder.setGeometry(QtCore.QRect(30, 80, 75, 20))
        self.gravity_folder.setObjectName("gravity_folder")

        # tableview
        self.tableview = QtWidgets.QTableView(self.central_widget)
        self.tableview.setGeometry(QtCore.QRect(190, 10, int(screen_size_x) - 225, int(screen_size_y) - 100))  # custom
        self.tableview.setObjectName("tableView")

        # open existing button
        self.open_existing_Btn = QtWidgets.QPushButton(self.options_pane)
        self.open_existing_Btn.setEnabled(False)
        self.open_existing_Btn.setGeometry(QtCore.QRect(20, 40, 113, 24))
        self.open_existing_Btn.setObjectName("open_existing_Btn")

        # create new button
        self.create_new_Btn = QtWidgets.QPushButton(self.options_pane)
        self.create_new_Btn.setEnabled(False)
        self.create_new_Btn.setGeometry(QtCore.QRect(20, 70, 113, 24))
        self.create_new_Btn.setObjectName("create_new_Btn")

        # add row button
        self.add_row = QtWidgets.QPushButton(self.options_pane)
        self.add_row.setEnabled(False)
        self.add_row.setGeometry(QtCore.QRect(20, 100, 113, 24))
        self.add_row.setObjectName("add_row")

        # delete row button
        self.del_row = QtWidgets.QPushButton(self.options_pane)
        self.del_row.setEnabled(False)
        self.del_row.setGeometry(QtCore.QRect(20, 130, 113, 24))
        self.del_row.setObjectName("del_row")

        # send to sign button
        self.send_to_sign = QtWidgets.QPushButton(self.options_pane)
        self.send_to_sign.setEnabled(False)
        self.send_to_sign.setGeometry(QtCore.QRect(20, 190, 113, 24))
        self.send_to_sign.setObjectName("send_to_sign")

        # save excel button
        self.save_btn = QtWidgets.QPushButton(self.options_pane)
        self.save_btn.setEnabled(False)
        self.save_btn.setGeometry(QtCore.QRect(20, 160, 113, 24))
        self.save_btn.setObjectName("save_btn")

        # reset button
        self.reset_btn = QtWidgets.QPushButton(self.options_pane)
        self.reset_btn.setEnabled(True)
        self.reset_btn.setGeometry(QtCore.QRect(20, 250, 113, 24))
        self.reset_btn.setObjectName("reset_btn")

        # send signed  button
        self.send_signed = QtWidgets.QPushButton(self.options_pane)
        self.send_signed.setEnabled(False)
        self.send_signed.setGeometry(QtCore.QRect(20, 220, 113, 24))
        self.send_signed.setObjectName("send_signed")

        self.import_csv = QtWidgets.QPushButton(self.options_pane)
        self.import_csv.setEnabled(False)
        self.import_csv.setGeometry(QtCore.QRect(20, 280, 113, 24))
        self.import_csv.setObjectName("send_signed")

        # Picture
        label = QtWidgets.QLabel(self.central_widget)
        pixmap = QtGui.QPixmap('O:\Field Services Division\Field Support Center\Project Acceptance\PA Excel Exterminator\imgs/OCULogo.jpg')
        label.setPixmap(pixmap)
        xpos = int(screen_size_x) - 225
        label.setGeometry(QtCore.QRect(int(xpos/2) - 100, 10, int(screen_size_x) - 225, int(screen_size_y) - 100))
        
        # setting names for the labels
        self.retranslateUi(MainWindow)
        self.isFirst = False

        # button clicked Functions              
        self.planfile_Btn.clicked.connect(lambda: search_clicked(self)) 
        self.open_existing_Btn.clicked.connect(lambda: init_table(self,"",""))
        self.add_row.clicked.connect(lambda: add_rows(self,self.tableview))
        self.del_row.clicked.connect(lambda: remove_row(self,self.tableview))
    
        self.send_signed.clicked.connect(lambda: send(self))
        self.import_csv.clicked.connect(lambda: import_from_book(self,self.tableview))
        
        self.create_new_Btn.clicked.connect(lambda: search_clicked(self))
        self.create_new_Btn.clicked.connect(lambda: create_new(self))
        self.send_to_sign.clicked.connect(lambda: pandas2word(self))
        self.send_to_sign.clicked.connect(lambda: self.send_signed.setEnabled(True))
        self.save_btn.clicked.connect(lambda: check_b4_save(self))
        self.reset_btn.clicked.connect(lambda: self.reset_all())
        self.work_entry_Btn.clicked.connect(lambda:check_work_area(self))

        self.open_existing_Btn.clicked.connect(lambda:self.reset_btn.setEnabled(False))
        self.open_existing_Btn.clicked.connect(lambda:self.work_entry_Btn.setEnabled(False))
        self.open_existing_Btn.clicked.connect(lambda:self.planfile_Btn.setEnabled(False))
        self.open_existing_Btn.clicked.connect(lambda:self.add_row.setEnabled(True))
        self.open_existing_Btn.clicked.connect(lambda:self.del_row.setEnabled(True))
        self.open_existing_Btn.clicked.connect(lambda:self.send_to_sign.setEnabled(True))
        self.open_existing_Btn.clicked.connect(lambda:self.save_btn.setEnabled(True))
        self.open_existing_Btn.clicked.connect(lambda:self.open_existing_Btn.setEnabled(False))

        self.development_checkB.stateChanged.connect(self.onStateChange)
        self.CIP_checkB.stateChanged.connect(self.onStateChange) 
        
        self.open_existing_Btn.clicked.connect(lambda:label.setHidden(True))
        self.create_new_Btn.clicked.connect(lambda:label.setHidden(True))

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Project Acceptance"))
        self.planfile_label.setText(_translate("MainWindow", "Plan File #"))
        self.planfile_Btn.setText(_translate("MainWindow", "Enter"))
        self.open_folder_label.setText(_translate("MainWindow", "Active Category"))
        self.pump_checkB.setText(_translate("MainWindow", "Pump"))
        self.pressure_checkB.setText(_translate("MainWindow", "Pressure"))
        self.gravity_checkB.setText(_translate("MainWindow", "Gravity"))
        self.options_Label.setText(_translate("MainWindow", "Options"))
        self.open_existing_Btn.setText(_translate("MainWindow", "Open Existing"))
        self.add_row.setText(_translate("MainWindow", "Add Row"))
        self.del_row.setText(_translate("MainWindow", "Delete Row"))
        self.create_new_Btn.setText(_translate("MainWindow", "Create New"))
        self.send_to_sign.setText(_translate("MainWindow", "Send for Signature"))
        self.save_btn.setText(_translate("MainWindow", "Save Changes"))
        self.reset_btn.setText(_translate("MainWindow", "Reset All Fields"))
        self.work_area_label.setText(_translate("MainWindow", "Work Area"))
        self.work_entry_Btn.setText(_translate("MainWindow", "Enter"))
        self.send_signed.setText(_translate("MainWindow", "Send Signed"))
        self.Dev_or_CIP_label.setText(_translate("MainWindow", "Development or CIP"))
        self.development_checkB.setText(_translate("MainWindow", "Development"))
        self.CIP_checkB.setText(_translate("MainWindow", "CIP"))
        self.import_csv.setText(_translate("MainWindow", "Import CSV from WB"))
        self.active_folder_label.setText(_translate("MainWindow", "Folders Found"))
        self.pump_folder.setText(_translate("MainWindow", "Pump"))
        self.pressure_folder.setText(_translate("MainWindow", "Pressure"))
        self.gravity_folder.setText(_translate("MainWindow", "Gravity"))

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
            self.send_signed.setEnabled(True)
            

        elif self.gravity_checkB.isChecked() and self.gravity_folder.isChecked():
            self.open_existing_Btn.setEnabled(True)
            self.create_new_Btn.setEnabled(False)
            self.send_signed.setEnabled(True)
            

        elif self.pressure_checkB.isChecked() and self.pressure_folder.isChecked():

            self.open_existing_Btn.setEnabled(True)
            self.create_new_Btn.setEnabled(False)
            self.send_signed.setEnabled(True)


        else:
            self.open_existing_Btn.setEnabled(False)
            self.create_new_Btn.setEnabled(True)
            self.send_to_sign.setEnabled(False)
            self.save_btn.setEnabled(False)
            self.add_row.setEnabled(False)
            self.del_row.setEnabled(False)
            

    
    def reset_all(self):
        self.development_checkB.setChecked(False)
        self.development_checkB.setEnabled(False)
        self.CIP_checkB.setChecked(False)
        self.CIP_checkB.setEnabled(False)
        self.work_entry.clear()
        self.planfile_entry.clear()
        self.save_btn.setEnabled(False)
        self.planfile_entry.setEnabled(False)
        self.send_to_sign.setEnabled(False)
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


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    app.setStyle('Fusion')
    MainWindow = QtWidgets.QMainWindow()
    ui = UI_Main_Window(MainWindow)
    ui.__init__(MainWindow)
    MainWindow.setWindowIcon(QtGui.QIcon('O:\Field Services Division\Field Support Center\Project Acceptance\PA Excel Exterminator\imgs/Logo.jpg'))

    MainWindow.show()
    sys.exit(app.exec())