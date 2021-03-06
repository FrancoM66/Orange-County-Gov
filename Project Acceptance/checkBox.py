from PyQt6.QtWidgets import QMessageBox
from PyQt6 import QtGui

def check_work_area(self):

    pump_string = {'Pump', 'pump', 'pumpstation', 'PumpStation', 'Pumpstation'}
    pressure_string = {"Pressure", "Pressurized","Pressurized Pipe", "Pipe", "pressure", "pressurized"}
    wastewater_string = {"Gravity", "gravity", "wastewater", "WasteWater"}

    work_entry_text = self.work_entry.text()

    if work_entry_text in pump_string:
        self.pump_checkB.setChecked(True)
        self.pressure_checkB.setChecked(False)
        self.gravity_checkB.setChecked(False)
        self.planfile_Btn.setEnabled(True)
        self.planfile_entry.setEnabled(True)
    elif work_entry_text in pressure_string:
        self.pump_checkB.setChecked(False)
        self.pressure_checkB.setChecked(True)
        self.gravity_checkB.setChecked(False)
        self.planfile_Btn.setEnabled(True)
        self.planfile_entry.setEnabled(True)
    elif work_entry_text in wastewater_string:
        self.pump_checkB.setChecked(False)
        self.gravity_checkB.setChecked(True)
        self.pressure_checkB.setChecked(False)
        self.planfile_Btn.setEnabled(True)
        self.planfile_entry.setEnabled(True)
    else:
        showError()
        self.planfile_Btn.setEnabled(False)
        self.planfile_entry.setEnabled(False)
        


def showError():
    msgBox = QMessageBox()
    msgBox.setWindowIcon(QtGui.QIcon('O:\Field Services Division\Field Support Center\Project Acceptance\PA Excel Exterminator\imgs/Logo.jpg'))
    msgBox.setText("Please Enter Valid Entry")
    msgBox.setWindowTitle("Error")
    msgBox.exec()
