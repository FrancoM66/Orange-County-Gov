from PyQt6 import QtGui
from PyQt6 import QtCore
from PyQt6 import QtWidgets

import pandas as pd
from PyQt6.QtCore import QAbstractTableModel, QTimer, Qt, pyqtSlot
from PyQt6.QtWidgets import QComboBox, QItemDelegate, QListWidget, QListWidgetItem, QMessageBox, QWidget


class PandasModel(QAbstractTableModel):
    def __init__(self, data):
        super().__init__()
        self._data = data

    def rowCount(self, index):
        return self._data.shape[0]

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
                    return QtGui.QColor(207,225,167)

                if (
                    (isinstance(value, str))
                    and value == "Rejected"
                ):
                    return QtGui.QColor(187,28,42)
                


    def setData(self, index, value, role):
        if role == Qt.ItemDataRole.EditRole:
            self._data.iloc[index.row(), index.column()] = value
            return True
        return False

    def headerData(self, col, orientation, role):
        if orientation == Qt.Orientation.Horizontal and role == Qt.ItemDataRole.DisplayRole:
            return self._data.columns[col]

    def flags(self, index):
        return Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsEditable


class PopupView(QWidget):
    def __init__(self, parent=None):
        super(PopupView, self).__init__(parent)
        self.setWindowFlags(Qt.Popup)
        self.move(QtGui.QCursor.pos())
        self.show()


class ItemDelegate(QItemDelegate):
    def __init__(self, parent):
        QItemDelegate.__init__(self, parent)

    def createEditor(self, parent, option, index):
        return PopupView(parent)


def init_table(self, table, concat, pump_checkB,gravity_checkB,pressure_checkB,cip_checkB,development_checkB):

    print("in excel.py")
    combo_box_options = ["Accepted","Rejected"]

    if pump_checkB.isChecked() and cip_checkB.isChecked():
        filename = "CIP Pump.csv"
    if pump_checkB.isChecked() and development_checkB.isChecked():
        filename = "Development Pump.csv"

    if gravity_checkB.isChecked() and cip_checkB.isChecked():
        filename = "CIP Gravity.csv"
    if gravity_checkB.isChecked() and development_checkB.isChecked():
        filename = "Development Gravity.csv"

    if pressure_checkB.isChecked() and cip_checkB.isChecked():
        filename = "CIP Pressurized.csv"
    if pressure_checkB.isChecked() and development_checkB.isChecked():
        filename = "Development Pressurized.csv"

    excel_filename = concat + "/Excel/" + filename
    print(excel_filename)
    df = pd.read_csv(excel_filename)
    if df.size == 0:
        return

    filtered_df = df.loc[:, ~df.columns.isin(["Notes", "Notes.1"])]

    self.model = PandasModel(df)

    #table.setModel(self.model)
    table.resizeRowsToContents()
    table.resizeColumnToContents(0)
    table.resizeColumnToContents(1)
    table.resizeColumnToContents(2)
    table.resizeColumnToContents(3)
    table.resizeColumnToContents(4)
    table.setColumnWidth(5, 300)
    table.resizeColumnToContents(6)
    table.setColumnWidth(7, 300)
    
    table.setWordWrap(True)

    table.setItemDelegate(ItemDelegate(table))

    table.show()

def showError():
    msgBox = QMessageBox()
    msgBox.setText("File Not Found")
    msgBox.setWindowTitle("Error")
    msgBox.exec()

