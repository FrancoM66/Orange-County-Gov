import os
from functions.checkpath import *
from PyQt6.QtWidgets import QMessageBox
from PyQt6 import QtWidgets
from functions.createfolder import *


def search_clicked(self):
    path = "O:\Field Services Division\Field Support Center\Project Acceptance"
    found = False
    workOrder = self.planfile_entry.text()
    if workOrder != "" and len(workOrder) >= 5:
        for root, subdir, files in os.walk(path):
            for d in subdir:
                if d.find(workOrder) != -1:
                    found = True
                    print(d)
                    print("Im HERE")
                    walk = d
                    self.concat = path + "/" + walk
                    isdir = os.path.isdir(self.concat)
                    if isdir:
                        pump_Found, pressure_Found, gravity_Found = check_path(self.concat)
                        print(str(pump_Found) + " " +
                              str(pressure_Found) + " " + str(gravity_Found))
                        self.development_checkB.setEnabled(True)
                        self.CIP_checkB.setEnabled(True)
                        if pump_Found == 1:
                            self.pump_folder.setChecked(True)
                            
                        if pressure_Found == 1:
                            self.pressure_folder.setChecked(True)
                            
                        if gravity_Found == 1:
                           self.gravity_folder.setChecked(True)

                    else:
                        pass
                
            break   
    else:
        showError()

    if not found:
        createNew(self) 


def showError():
    msgBox = QMessageBox()
    msgBox.setText("Please Enter Valid Entry")
    msgBox.setWindowTitle("Error")
    msgBox.exec()

def createNew(self):
    msgBox = QMessageBox()
    msgBox.setText("Folder not found. Create new folder with this planfile?")
    msgBox.setWindowTitle("Create new folder")
    msgBox.setStandardButtons(
        QtWidgets.QMessageBox.StandardButton.Ok | QtWidgets.QMessageBox.StandardButton.Cancel
    )
    response = msgBox.exec()
    print(response)
    if response == 1024:
        create_planfile_folder(self)
        pass
    else:
        pass
