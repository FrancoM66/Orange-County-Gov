from fileinput import filename
from PyQt6 import QtGui
from PyQt6 import QtCore
from PyQt6 import QtWidgets

import pandas as pd
from PyQt6.QtCore import QAbstractTableModel, QModelIndex, QSortFilterProxyModel, QTimer, Qt, pyqtSlot
from PyQt6.QtWidgets import QComboBox, QItemDelegate, QListWidget, QListWidgetItem, QMessageBox, QTableView, QWidget
import configparser
import json

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

def filename_checker(self, concat, pump_checkB, gravity_checkB, pressure_checkB, cip_checkB, development_checkB):

    if pump_checkB.isChecked() and cip_checkB.isChecked():
        self.filename = "CIP Pump.csv"
        self.code = 1

    if pump_checkB.isChecked() and development_checkB.isChecked():
        self.filename = "Development Pump.csv"
        self.code = 2

    if gravity_checkB.isChecked() and cip_checkB.isChecked():
        self.filename = "CIP Gravity.csv"
        self.code = 3

    if gravity_checkB.isChecked() and development_checkB.isChecked():
        self.filename = "Development Gravity.csv"
        self.code = 4

    if pressure_checkB.isChecked() and cip_checkB.isChecked():
        self.filename = "CIP Pressurized.csv"
        self.code = 5

    if pressure_checkB.isChecked() and development_checkB.isChecked():
        self.filename = "Development Pressurized.csv"
        self.code = 6

    self.excel_filename = concat + "/Excel/" + self.filename
    

# Initialized the table
def init_table(self, table, concat, pump_checkB, gravity_checkB, pressure_checkB, cip_checkB, development_checkB,template, variation):

    filename_checker(self, concat, pump_checkB, gravity_checkB,
                     pressure_checkB, cip_checkB, development_checkB)

    if template!= "":
        self.excel_filename = template
        self.variation = variation

    print("in excel.py")
    print(self.excel_filename)
    self.df = pd.read_csv(self.excel_filename)
    if self.df.size == 0:
        return

    # filtered_df = df.loc[:, ~df.columns.isin(["Notes", "Notes.1"])]
    set_delegates(self, table)

    self.model = PandasModel(self.df)

    table.setModel(self.model)
    table.resizeRowsToContents()
    table.setWordWrap(True)

    self.setter = table
    
    table.show()

def add_rows(self, table):
    df2 = {'Warranty': ''}

    self.df = self.df.append(df2, ignore_index = True)
    self.model = PandasModel(self.df)

    set_delegates(self, table)

    table.setModel(self.model)
    table.show()

def remove_row(self, table):
    x = table.selectionModel().selectedIndexes()
    print(x)


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
    category_delegate.setItems(pump_value_list)
   
    if self.code == 1:
        table.setItemDelegateForColumn(0, category_delegate)
        table.setItemDelegateForColumn(4, options_delegate)
        table.setItemDelegateForColumn(6, options_delegate)

    if self.code == 2:
        table.setItemDelegateForColumn(0, category_delegate)
        table.setItemDelegateForColumn(4, options_delegate)
        table.setItemDelegateForColumn(6, options_delegate)

    if self.code == 3:
        table.setItemDelegateForColumn(4, structure_delegate)
        table.setItemDelegateForColumn(5, action_delegate)
        table.setItemDelegateForColumn(7, approved_cctv_delegate)
        table.setItemDelegateForColumn(8, options_delegate)
        table.setItemDelegateForColumn(10, reviewer_delegate)
        table.setItemDelegateForColumn(11, corrective_delegate)
        table.setItemDelegateForColumn(13, options_delegate)
        table.setItemDelegateForColumn(14, vendor)
        table.setItemDelegateForColumn(15, reviewer_delegate)
        table.setItemDelegateForColumn(16, corrective_delegate)

    if self.code == 4:
        table.setItemDelegateForColumn(2, approved_cctv_delegate)
        table.setItemDelegateForColumn(3, options_delegate)
        table.setItemDelegateForColumn(5, reviewer_delegate)
        table.setItemDelegateForColumn(7, options_delegate)
        table.setItemDelegateForColumn(8, vendor)
        table.setItemDelegateForColumn(9, reviewer_delegate)

    if self.code == 5:
        table.setItemDelegateForColumn(0, category_delegate)
        table.setItemDelegateForColumn(4, structure_delegate)
        table.setItemDelegateForColumn(6, structure_delegate)

    if self.code == 6:
        table.setItemDelegateForColumn(0, category_delegate)
        table.setItemDelegateForColumn(3, structure_delegate)
        table.setItemDelegateForColumn(4, options_delegate)
        table.setItemDelegateForColumn(6, options_delegate)


def exportToExcel(self):
    columnHeaders = []
    

    # create column header list
    for j in range(self.setter.model().columnCount()):
        columnHeaders.append(self.setter.model().headerData(
            j, Qt.Orientation.Horizontal, Qt.ItemDataRole.DisplayRole))

    dfnew = pd.DataFrame(columns=columnHeaders)

    # create dataframe object recordset
    for row in range(self.setter.model().rowCount()):
        for col in range(self.setter.model().columnCount()):
            dfnew.at[row, columnHeaders[col]
                     ] = self.setter.model().index(row, col).data()
    # TODO: Change this from hard coding
    dfnew.to_csv(self.concat + "/Excel" + self.variation, index=False)
    print('Excel file exported')


def showError():
    msgBox = QMessageBox()
    msgBox.setText("File Not Found")
    msgBox.setWindowTitle("Error")
    msgBox.exec()


def check_b4_save(self):

    msgBox = QMessageBox()
    msgBox.setText("All changes will be saved")
    msgBox.setWindowTitle("Save File")
    msgBox.setStandardButtons(
        QtWidgets.QMessageBox.StandardButton.Ok | QtWidgets.QMessageBox.StandardButton.Cancel
    )
    response = msgBox.exec()
    print(response)
    if response == 1024:
        exportToExcel(self)
    else:
        pass
