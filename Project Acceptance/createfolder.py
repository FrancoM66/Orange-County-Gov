import os
import main
import excel
from PyQt6.QtWidgets import QMessageBox

def create_new(self, concat):
        template = "O:\Field Services Division\Field Support Center\Project Acceptance\PA Excel Exterminator\Templates"
        parent_dir = concat
        print("In create New " + parent_dir)

        if self.pump_checkB.isChecked():
            if self.CIP_checkB.isChecked():
                variation = "/Cip Pump.csv"
                csv_tranfer = template + variation
            elif self.development_checkB.isChecked():
                variation = "/Development Pump.csv"
                csv_tranfer = template + variation
            self.directoryCode = "/Pump Station"
            path = parent_dir+ self.directoryCode
            os.mkdir(path) 
            excel.init_table(
            self, self.tableview, concat, self.pump_checkB, self.gravity_checkB, self.pressure_checkB, self.CIP_checkB, self.development_checkB, csv_tranfer, variation)
            print("Directory '% s' created" % self.directoryCode) 
            print(path)
            self.reset_btn.setEnabled(False)
            self.work_entry_Btn.setEnabled(False)
            self.planfile_Btn.setEnabled(False)
            self.add_row.setEnabled(True)
            self.create_pdf.setEnabled(True)
            self.save_btn.setEnabled(True)
            create_successful(self)
 

        elif self.gravity_checkB.isChecked():
            if self.CIP_checkB.isChecked():
                variation = "/Cip Gravity.csv"
                csv_tranfer = template + variation
            elif self.development_checkB.isChecked():
                variation = "/Development Gravity.csv"
                csv_tranfer = template + variation
            self.directoryCode = "/Wastewater"
            path = parent_dir+ self.directoryCode
            print(path)
            os.mkdir(path) 
            excel.init_table(
            self, self.tableview, concat, self.pump_checkB, self.gravity_checkB, self.pressure_checkB, self.CIP_checkB, self.development_checkB, csv_tranfer, variation)
            print("Directory '% s' created" % self.directoryCode) 
            self.reset_btn.setEnabled(False)
            self.work_entry_Btn.setEnabled(False)
            self.planfile_Btn.setEnabled(False)
            self.add_row.setEnabled(True)
            self.create_pdf.setEnabled(True)
            self.save_btn.setEnabled(True)
            create_successful(self)

        elif self.pressure_checkB.isChecked():
            if self.CIP_checkB.isChecked():
                variation = "/Cip Pressure.csv"
                csv_tranfer = template + variation
            elif self.development_checkB.isChecked():
                variation = "/Development Pressure.csv"
            self.directoryCode = "/Pressurized Pipe"
            path = parent_dir+ self.directoryCode
            os.mkdir(path) 
            excel.init_table(
            self, self.tableview, concat, self.pump_checkB, self.gravity_checkB, self.pressure_checkB, self.CIP_checkB, self.development_checkB, csv_tranfer, variation)
            print("Directory '% s' created" % self.directoryCode) 
            self.reset_btn.setEnabled(False)
            self.work_entry_Btn.setEnabled(False)
            self.planfile_Btn.setEnabled(False)
            self.add_row.setEnabled(True)
            self.create_pdf.setEnabled(True)
            self.save_btn.setEnabled(True)
            create_successful(self)

    
def create_planfile_folder(self,workOrder):
    planfile_folder = "O:\Field Services Division\Field Support Center\Project Acceptance"
    fill_planfile(self, workOrder)
    pass

def create_successful(self):
    msgBox = QMessageBox()
    msgBox.setText("You created " + self.directoryCode)
    msgBox.setWindowTitle("Successful")
    self.create_new_Btn.setEnabled(False)
    self.open_existing_Btn.setEnabled(True)
    msgBox.exec()

def fill_planfile(self, workOrder):
    msgBox = QMessageBox()
    msgBox.setText(workOrder)
    msgBox.setWindowTitle("Successful")
    msgBox.exec()