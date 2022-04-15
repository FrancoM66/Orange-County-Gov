import os
from PyQt6.QtWidgets import QMessageBox,QInputDialog,QWidget
from PyQt6 import QtWidgets
from checkpath import * 
from PyQt6 import QtGui

class App(QWidget):
    global acceptedorrejected
    global walkthroughorwarranty
    global workorder

    def __init__(self):
        super().__init__()
        self.title = 'popup'
        self.left = 10
        self.top = 10
        self.width = 640
        self.height = 480
        self.setWindowIcon(QtGui.QIcon('O:\Field Services Division\Field Support Center\Project Acceptance\PA Excel Exterminator\imgs/Logo.jpg'))
        self.initUI()
    
    def initUI(self):
        self.getText()
        

    def getText(self):
        self.projectname, okPressed = QInputDialog.getText(self, " ","Project Name:",  QtWidgets.QLineEdit.EchoMode.Normal, "")
        if okPressed and self.projectname != '':
            print(self.projectname)

        self.ocNum, okPressed = QInputDialog.getText(self, " ","Project OC Number:",  QtWidgets.QLineEdit.EchoMode.Normal, "")
        if okPressed and self.ocNum != '':
            print(self.ocNum)

    def center(self):
        qr = self.frameGeometry()
        cp = QtGui.QGuiApplication.primaryScreen().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

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
                    self.mend = self.concat
                    isdir = os.path.isdir(self.concat)
                    if isdir:
                        pump_Found, pressure_Found, gravity_Found, excel = check_path(self.concat)
                        print(str(pump_Found) + " " + str(pressure_Found) + " " + str(gravity_Found))
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

    if not found and workOrder != "" and len(workOrder) >= 5:
        createNew(self) 


def showError():
    msgBox = QMessageBox()
    msgBox.setWindowIcon(QtGui.QIcon('O:\Field Services Division\Field Support Center\Project Acceptance\PA Excel Exterminator\imgs/Logo.jpg'))
    msgBox.setText("Please Enter Valid Entry")
    msgBox.setWindowTitle("Error")
    msgBox.exec()

def createNew(self):
    workOrder = self.planfile_entry.text()
    self.isFirst = True
    print("in createNew: " + str(self.isFirst))
    msgBox = QMessageBox()
    msgBox.setWindowIcon(QtGui.QIcon('O:\Field Services Division\Field Support Center\Project Acceptance\PA Excel Exterminator\imgs/Logo.jpg'))
    msgBox.setText("Folder not found. Create new folder with this planfile?")
    msgBox.setWindowTitle("Create new folder")
    msgBox.setStandardButtons(QtWidgets.QMessageBox.StandardButton.Ok | QtWidgets.QMessageBox.StandardButton.Cancel)
    response = msgBox.exec()

    print(response)
    if response == 1024:
        self.getchoice = App()
        create_planfile_folder(self, workOrder, self.getchoice.projectname)
        create_info(workOrder,self.getchoice,"")
        pass
    else:
        pass

def create_planfile_folder(self,workOrder,projectname):
    planfile_folder = "O:\Field Services Division\Field Support Center\Project Acceptance"
    self.path =  planfile_folder + "/" + workOrder + " - " + projectname
    os.mkdir(self.path)
    makeXL = self.path + "/Excel"
    os.mkdir(makeXL)
    self.development_checkB.setEnabled(True)
    self.CIP_checkB.setEnabled(True)


def create_info(workorder,getchoice,inspector):
    path = r"O:\Field Services Division\Field Support Center\Project Acceptance\PA Excel Exterminator\info"
    filename = workorder + ".txt"
    file_exists = os.path.exists(path+ "/" +filename)
    print(file_exists)
    if file_exists == False:
        f = open(path + "/" +filename, "w+")
        f.write(getchoice.projectname + "\n")
        f.write(getchoice.ocNum + "\n")
        f.close()
    if file_exists == True:
        

        file_name = open(path + "/" + filename,"r")
        Counter = 0
        
        # Reading from file
        Content = file_name.read()
        CoList = Content.split("\n")
        
        for i in CoList:
            if i:
                Counter += 1
                print(Counter)
        
        lines = open(path + "/" + filename, 'r').readlines()
        print(lines)
        if Counter > 2:
            lines[-1] = inspector
            open(path + "/" + filename, 'w').writelines(lines)
        if Counter == 2:
            f = open(path + "/" +filename, "a+")
            f.write(inspector)
            f.close()
        
        