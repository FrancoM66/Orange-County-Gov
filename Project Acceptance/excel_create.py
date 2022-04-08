from codecs import ignore_errors
from PyQt6 import QtGui
from PyQt6 import QtWidgets
from PyQt6.QtCore import QAbstractTableModel, QModelIndex,Qt
from PyQt6.QtWidgets import QComboBox, QItemDelegate,QMessageBox,QInputDialog,QWidget
import pandas as pd
import configparser
import json
from word_create import *
from mail import *
from docx_create import *
from search import *

# Creates the model for the QTableView
class PandasModel(QAbstractTableModel):
    def __init__(self, data):
        super().__init__()
        self._data = data
        self.__empty_row_count = 0
        self.__total_row_count = 0  # Max visible rows.

    def rowCount(self, index=QModelIndex()):                    
        return  0 if index.isValid() else len(self._data)

    def columnCount(self, parnet=None):
        return self._data.shape[1]

    def data(self, index, role):
        if index.isValid():
            if role == Qt.ItemDataRole.DisplayRole or role == Qt.ItemDataRole.EditRole:
                value = self._data.iloc[index.row(), index.column()]
                return str(value)

        if index.isValid():
            if role == Qt.ItemDataRole.BackgroundRole:
                value = self._data.iloc[index.row(), index.column()]
                if (
                    (isinstance(value, str))
                    and value == "Accepted"
                ):
                    return QtGui.QColor(207, 225, 167)

                if (
                    (isinstance(value, str))
                    and value == "Rejected"
                ):
                    return QtGui.QColor(187, 28, 42)

                if (
                    (isinstance(value, str))
                    and value == "Removed"
                ):
                    return QtGui.QColor(255, 191, 0)

    def setData(self, index, value, role):
        if role == Qt.ItemDataRole.EditRole:
            self._data.iloc[index.row(), index.column()] = value
            return True
        return False

    def emptyRowCount(self):
        return self.__empty_row_count

    def headerData(self, col, orientation, role):
        if orientation == Qt.Orientation.Horizontal and role == Qt.ItemDataRole.DisplayRole:
            return self._data.columns[col]

    def flags(self, index):
        return Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsEditable

# Creates the ComboBoxes for the QTableView
class ComboBoxDelegate(QItemDelegate):
    
    def __init__(self, parent=None):
        super(ComboBoxDelegate, self).__init__(parent)
        self.items = []
        

    def setItems(self, items):
        self.items = items

    def createEditor(self, widget, option, index):
        editor = QComboBox(widget)
        editor.addItems(self.items)
        return editor

    def setEditorData(self, editor, index):
        value = index.model().data(index, Qt.ItemDataRole.EditRole)

        if value:
            editor.setCurrentText(str(value))

    def setModelData(self, editor, model, index):
        model.setData(index, editor.currentText(), Qt.ItemDataRole.EditRole)

    def updateEditorGeometry(self, editor, option, index):
        editor.setGeometry(option.rect)

structure_delegate = ComboBoxDelegate()
category_delegate = ComboBoxDelegate()
category_pump_delegate = ComboBoxDelegate()
approved_cctv_delegate = ComboBoxDelegate()
vendor = ComboBoxDelegate()
reviewer_delegate = ComboBoxDelegate()
delegate = ComboBoxDelegate()
delegate.setItems(["Accepted", "Rejected"])
action_delegate = ComboBoxDelegate()
options_delegate = ComboBoxDelegate()
options_delegate.setItems(["Accepted", "Rejected", "Removed"])
corrective_delegate = ComboBoxDelegate()
empty_delegate = ComboBoxDelegate()

def filename_checker(self):

        if self.pump_checkB.isChecked() and self.CIP_checkB.isChecked():
            self.filename = "CIP Pump.csv"
            self.code = 1
            self.directory_code = self.concat+"/Pump-Station"
            self.area = "Pump-Station"
            self.cip_dev = "CIP"

        if self.pump_checkB.isChecked() and self.development_checkB.isChecked():
            self.filename = "Development Pump.csv"
            self.code = 2
            self.directory_code = self.concat+"/Pump-Station"
            self.area = "Pump-Station"
            self.cip_dev = "Dev"

        if self.gravity_checkB.isChecked() and self.CIP_checkB.isChecked():
            self.filename = "Cip Gravity.csv"
            self.code = 3
            self.directory_code = self.concat+"/Wastewater"
            self.area = "Gravity"
            self.cip_dev = "CIP"

        if self.gravity_checkB.isChecked() and self.development_checkB.isChecked():
            self.filename = "Development Gravity.csv"
            self.code = 4
            self.directory_code = self.concat+"/Wastewater"
            self.area = "Gravity"
            self.cip_dev = "Dev"

        if self.pressure_checkB.isChecked() and self.CIP_checkB.isChecked():
            self.filename = "CIP Pressure.csv"
            self.code = 5
            self.directory_code = self.concat+"/Pressurized-Pipe"
            self.area = "Pressurized-Pipe"
            self.cip_dev = "CIP"

        if self.pressure_checkB.isChecked() and self.development_checkB.isChecked():
            self.filename = "Development Pressure.csv"
            self.code = 6
            self.directory_code = self.concat+"/Pressurized-Pipe"
            self.area = "Pressurized-Pipe"
            self.cip_dev = "Dev"

        self.excel_filename = self.concat + "/Excel/" + self.filename

class App(QWidget):
    
    def __init__(self, workOrder):
        super().__init__()
        self.title = 'popup'
        self.left = 10
        self.top = 10
        self.width = 640
        self.height = 480
        self.setWindowIcon(QtGui.QIcon('O:\Field Services Division\Field Support Center\Project Acceptance\PA Excel Exterminator\imgs/Logo.jpg'))
        self.workOrder = workOrder
        print(self.workOrder + " Inside the init")
        self.initUI()
    
    def initUI(self):
        self.getChoice()
        self.getText()

    def getChoice(self):

        config = configparser.RawConfigParser()
        config.read("O:\Field Services Division\Field Support Center\Project Acceptance\PA Excel Exterminator\config\emailCC.properties")
        inspectors = config.get("INSPECTORS", "inspectorList")
        inspectorList = json.loads(inspectors)

        chiefInspector = config.get("INSPECTORS", "chief_inspectors")
        chiefInspectorList = json.loads(chiefInspector)

        accepted_rejected = ("","Acceptance","Rejection")
        walkthrough_warranty = ("Walkthrough","Warranty")
        self.item2, okPressed = QInputDialog.getItem(self, " ","Walkthrough or Warranty:", walkthrough_warranty, 0, False)
        self.item, okPressed = QInputDialog.getItem(self, " ","Accepted or Rejected:", accepted_rejected, 0, False)
        self.inspector, okPressed = QInputDialog.getItem(self, " ","Choose an inspector:", inspectorList, 0, False)
        self.chiefInspector, okPressed = QInputDialog.getItem(self, " ","Choose an inspector:", chiefInspectorList, 0, False)
        create_info(self.workOrder,"",self.inspector)

    def getText(self):
        self.text, okPressed = QInputDialog.getText(self, " ","Please Enter Inspection Date mm-dd-yyyy:",  QtWidgets.QLineEdit.EchoMode.Normal, "")

        self.text2, okPressed = QInputDialog.getText(self, " ","Please Enter Refrence Work Order:",  QtWidgets.QLineEdit.EchoMode.Normal, "")

    

    def center(self):
        qr = self.frameGeometry()
        cp = QtGui.QGuiApplication.primaryScreen().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

def send(self):
    filename_checker(self)
    workOrder_pf = self.planfile_entry.text()
    mail_signed(self,self.directory_code, workOrder_pf)

# Initialized the table
def init_table(self,template, variation):

    filename_checker(self)

    if template!= "":
        self.excel_filename = template
        self.variation1 = variation
    else:
        self.variation1 = "/"+ self.filename
    
    self.df = pd.read_csv(self.excel_filename)
    self.df = self.df.fillna("N/A")
    
    if self.df.size == 0:
        return

    
    set_delegates(self, self.tableview)

    self.model = PandasModel(self.df)

    self.tableview.setModel(self.model)
    self.tableview.setWordWrap(True)
    self.tableview.horizontalHeader().setStretchLastSection(True)
    
    self.tableview.show()


def add_rows(self, table):
    df2 = {'Warranty': ''}
    self.df = self.df.append(df2, ignore_index = True)
    self.df = self.df.fillna("N/A")

    self.model = PandasModel(self.df)

    set_delegates(self, table)

    table.setModel(self.model)
    
    table.show()

def import_from_book(self,table):
    file_name = QtWidgets.QFileDialog.getOpenFileName(None, "Select Directory")
    if(file_name == ""):
        return
    else:

        df_to_merge = pd.read_csv(file_name[0],skiprows=9)
        
        for x in range(df_to_merge.shape[1]):
            self.df["Location"] = pd.concat([df_to_merge["test"].dropna(),self.df["Location"]], ignore_index=True)

        self.model = PandasModel(self.df)

        set_delegates(self, table)

        table.setModel(self.model)
        
        table.show()


def remove_row(self, table):
    x = table.selectionModel().currentIndex()
    NewIndex = x.row()
    new_model = self.df.drop(self.df.index[NewIndex])
    self.model = PandasModel(new_model)

    set_delegates(self, table)

    table.setModel(self.model)
    self.df = new_model
    table.show()
    
def set_delegates(self, table):

    config = configparser.RawConfigParser()
    config.read(
        "O:\Field Services Division\Field Support Center\Project Acceptance\PA Excel Exterminator\config\delegates.properties")

    approved_cctv = config.get("GRAVITY", "approved_cctv")
    approved_cctv_list = json.loads(approved_cctv)
    approved_cctv_delegate.setItems(approved_cctv_list)

    vendor_surveyor = config.get("GRAVITY", "oc_surveyor")
    vendor_surveyor_list = json.loads(vendor_surveyor)
    vendor.setItems(vendor_surveyor_list)

    reviewer = config.get("GRAVITY", "reviewer")
    reviewer_list = json.loads(reviewer)
    reviewer_delegate.setItems(reviewer_list)

    gravity_structure = config.get("GRAVITY", "structure")
    gravity_structure_list = json.loads(gravity_structure)
    structure_delegate.setItems(gravity_structure_list)

    action = config.get("GRAVITY", "action")
    action_list = json.loads(action)
    action_delegate.setItems(action_list)

    corrective = config.get("GRAVITY", "corrective")
    corrective_list = json.loads(corrective)
    corrective_delegate.setItems(corrective_list)

    pressure_structure_vals = config.get("PRESSURE", "structure")
    pressure_structure_vals_list = json.loads(pressure_structure_vals)
    structure_delegate.setItems(pressure_structure_vals_list)

    pressure_category_val = config.get("PRESSURE", "category")
    category_list = json.loads(pressure_category_val)
    category_delegate.setItems(category_list)

    pump_values = config.get("PUMP", "category")
    pump_value_list = json.loads(pump_values)
    category_pump_delegate.setItems(pump_value_list)

    if self.code == 1:
        table.setItemDelegateForColumn(0, category_pump_delegate)
        table.setItemDelegateForColumn(4, options_delegate)
        table.setItemDelegateForColumn(6, options_delegate)

    if self.code == 2:
        table.setItemDelegateForColumn(0, category_pump_delegate)
        table.setItemDelegateForColumn(4, options_delegate)
        table.setItemDelegateForColumn(6, options_delegate)

    if self.code == 3:
        table.setItemDelegateForColumn(4, structure_delegate)
        table.setItemDelegateForColumn(5, action_delegate)
        table.setItemDelegateForColumn(6, approved_cctv_delegate)
        table.setItemDelegateForColumn(7, options_delegate)
        table.setItemDelegateForColumn(8, vendor)
        table.setItemDelegateForColumn(9, reviewer_delegate)
        table.setItemDelegateForColumn(10, corrective_delegate)
        table.setItemDelegateForColumn(12, options_delegate)
        table.setItemDelegateForColumn(14, reviewer_delegate)
        table.setItemDelegateForColumn(15, corrective_delegate)

    if self.code == 4:
        table.setItemDelegateForColumn(2, approved_cctv_delegate)
        table.setItemDelegateForColumn(3, options_delegate)
        table.setItemDelegateForColumn(5, reviewer_delegate)
        table.setItemDelegateForColumn(7, options_delegate)
        table.setItemDelegateForColumn(8, vendor)
        table.setItemDelegateForColumn(9, reviewer_delegate)

    if self.code == 5:
        table.setItemDelegateForColumn(0, category_delegate)
        table.setItemDelegateForColumn(5, structure_delegate)
        table.setItemDelegateForColumn(6, options_delegate)
        table.setItemDelegateForColumn(8, options_delegate)

    if self.code == 6:
        table.setItemDelegateForColumn(0, category_delegate)
        table.setItemDelegateForColumn(5, structure_delegate)
        table.setItemDelegateForColumn(6, options_delegate)
        table.setItemDelegateForColumn(8, options_delegate)


def exportToExcel(self, table):
    columnHeaders = []
    setter = table
    # create column header list
    for j in range(setter.model().columnCount()):
        columnHeaders.append(setter.model().headerData(
            j, Qt.Orientation.Horizontal, Qt.ItemDataRole.DisplayRole))

    dfnew = pd.DataFrame(columns=columnHeaders)

    # create dataframe object recordset
    for row in range(setter.model().rowCount()):
        for col in range(setter.model().columnCount()):
            dfnew.at[row, columnHeaders[col]] = setter.model().index(row, col).data()
    
    dfnew.to_csv(self.concat + "/Excel" + self.variation1, index=False)


def pandas2word(self):
    workOrder_pf = self.planfile_entry.text()
    self.getchoice = App(workOrder_pf)
    
    if self.getchoice.item == "Acceptance" and self.getchoice.item2 == "Walkthrough":
        found = self.df[self.df['Walkthrough'].str.contains('Rejected')]
        found2 = len(self.df[self.df['Walkthrough'] == 'N/A']) 
        if len(found) == 0 and found2 == 0:
            filtered_df = variation(self)
            location = self.directory_code 
            create_word(self, filtered_df, self.area, self.cip_dev, location,self.getchoice.text, self.getchoice.item2, self.getchoice.item,self.getchoice.chiefInspector)
            acceptance_no_deficiencies(self, self.area, self.cip_dev, location,self.getchoice.text, self.getchoice.item2, self.getchoice.item,self.getchoice.text2, workOrder_pf)
            
        else:
            print("got u again")
            error_acceptance()

    if self.getchoice.item == "Acceptance" and self.getchoice.item2 == "Warranty":
        found = self.df[self.df['Warranty'].str.contains('Rejected')]
        found2 = len(self.df[self.df['Warranty'] == 'N/A']) 
        if len(found) == 0 and found2 == 0:
            filtered_df = variation(self)
            location = self.directory_code 
            create_word(self, filtered_df, self.area, self.cip_dev, location,self.getchoice.text, self.getchoice.item2, self.getchoice.item,self.getchoice.chiefInspector)
            acceptance_no_deficiencies(self, self.area, self.cip_dev, location,self.getchoice.text, self.getchoice.item2, self.getchoice.item,self.getchoice.text2, workOrder_pf)
        else:
            print("got u again")
            error_acceptance()
            

    if self.getchoice.item == "Rejection" and self.getchoice.item2 == "Walkthrough":
        found2 = len(self.df[self.df['Walkthrough'] == 'N/A']) 
        if found2 == 0:
            filtered_df = variation(self)
            location = self.directory_code 
            create_word(self, filtered_df, self.area, self.cip_dev, location,self.getchoice.text, self.getchoice.item2, self.getchoice.item,self.getchoice.chiefInspector)
            rejected_word(self, self.area, self.cip_dev, location,self.getchoice.text, self.getchoice.item2, self.getchoice.item,self.getchoice.text2, workOrder_pf)
        else:
                print("got u again")
                error_rejection()
                pass

    if self.getchoice.item == "Rejection" and self.getchoice.item2 == "Warranty":
        found2 = len(self.df[self.df['Warranty'] == 'N/A']) 
        if found2 == 0:
            filtered_df = variation(self)
            location = self.directory_code 
            create_word(self, filtered_df, self.area, self.cip_dev, location,self.getchoice.text, self.getchoice.item2, self.getchoice.item,self.getchoice.chiefInspector)
            rejected_word(self, self.area, self.cip_dev, location,self.getchoice.text, self.getchoice.item2, self.getchoice.item,self.getchoice.text2, workOrder_pf)
        else:
                print("got u again")
                error_rejection()

                pass


def variation(self):
    
    if self.code == 1:
        if self.getchoice.item2 == "Walkthrough":
            filtered_df = self.df.loc[:, ~self.df.columns.isin([ "Notes.1", "Warranty"])]
            return filtered_df
        if self.getchoice.item2 == "Warranty":
            filtered_df = self.df.loc[:, ~self.df.columns.isin(["Walkthrough", "Notes"])]
            return filtered_df

    if self.code == 2:
        if self.getchoice.item2 == "Walkthrough":
            filtered_df = self.df.loc[:, ~self.df.columns.isin([ "Notes.1", "Warranty"])]
            return filtered_df
        if self.getchoice.item2 == "Warranty":
            filtered_df = self.df.loc[:, ~self.df.columns.isin(["Walkthrough", "Notes"])]
            return filtered_df
    
    if self.code == 3:
        if self.getchoice.item2 == "Walkthrough":
            filtered_df = self.df.loc[:, ~self.df.columns.isin(["Notes.1", "Warranty","Package", "Subdivision", "Reviewer","Reviewer.1", "OC Surveyor", "Corrective Action", "Corrective Action.1","Action Recommendation","Approved CCTV Vendor", "Vendor Surveyor"])]
            return filtered_df
        if self.getchoice.item2 == "Warranty":
            filtered_df = self.df.loc[:, ~self.df.columns.isin(["Walkthrough", "Notes","Package", "Subdivision", "Reviewer","Reviewer.1", "OC Surveyor", "Corrective Action", "Corrective Action.1","Action Recommendation","Approved CCTV Vendor", "Vendor Surveyor"])]
            return filtered_df

    if self.code == 4:
        if self.getchoice.item2 == "Walkthrough":
            filtered_df = self.df.loc[:, ~self.df.columns.isin(["OC Surveyor", "Notes.1", "Warranty", "Reviewer.1"])]
            return filtered_df
        if self.getchoice.item2 == "Warranty":
            filtered_df = self.df.loc[:, ~self.df.columns.isin(["Walkthrough", "Notes", "Vendor Surveyor", "Reviewer"])]
            return filtered_df

    if self.code == 5:
        if self.getchoice.item2 == "Walkthrough":
            filtered_df = self.df.loc[:, ~self.df.columns.isin(["Notes.1", "Warranty"])]
            return filtered_df
        if self.getchoice.item2 == "Warranty":
            filtered_df = self.df.loc[:, ~self.df.columns.isin(["Walkthrough", "Notes"])]
            return filtered_df

    if self.code == 6:
        if self.getchoice.item2 == "Walkthrough":
            filtered_df = self.df.loc[:, ~self.df.columns.isin(["Notes.1", "Warranty"])]
            return filtered_df
        if self.getchoice.item2 == "Warranty":
            filtered_df = self.df.loc[:, ~self.df.columns.isin(["Walkthrough", "Notes"])]
            return filtered_df
    

def showError():
    msgBox = QMessageBox()
    msgBox.setWindowIcon(QtGui.QIcon('O:\Field Services Division\Field Support Center\Project Acceptance\PA Excel Exterminator\imgs/Logo.jpg'))
    msgBox.setText("File Not Found")
    msgBox.setWindowTitle("Error")
    msgBox.exec()

def error_acceptance():
    msgBox = QMessageBox()
    msgBox.setWindowIcon(QtGui.QIcon('O:\Field Services Division\Field Support Center\Project Acceptance\PA Excel Exterminator\imgs/Logo.jpg'))
    msgBox.setText("Rejection or empty space found when trying to send Acceptance Letter")
    msgBox.setWindowTitle("Error")
    msgBox.exec()

def error_rejection():
    msgBox = QMessageBox()
    msgBox.setWindowIcon(QtGui.QIcon('O:\Field Services Division\Field Support Center\Project Acceptance\PA Excel Exterminator\imgs/Logo.jpg'))
    msgBox.setText(" Empty space found when trying to send Rejection Letter")
    msgBox.setWindowTitle("Error")
    msgBox.exec()



def check_b4_save(self):

    msgBox = QMessageBox()
    msgBox.setText("All changes will be saved")
    msgBox.setWindowIcon(QtGui.QIcon('O:\Field Services Division\Field Support Center\Project Acceptance\PA Excel Exterminator\imgs/Logo.jpg'))
    msgBox.setWindowTitle("Save File")
    msgBox.setStandardButtons(
        QtWidgets.QMessageBox.StandardButton.Ok | QtWidgets.QMessageBox.StandardButton.Cancel
    )
    response = msgBox.exec()
    print(response)
    if response == 1024:
        exportToExcel(self, self.tableview)
    else:
        pass
