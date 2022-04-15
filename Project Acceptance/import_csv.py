from PyQt6 import QtGui
from PyQt6 import QtWidgets
from PyQt6.QtCore import QAbstractTableModel, QModelIndex,Qt
from PyQt6.QtWidgets import QComboBox, QItemDelegate,QMessageBox,QInputDialog,QWidget
import pandas as pd

def import_from_book(self):
    file_name = QtWidgets.QFileDialog.getOpenFileName(None, "Select Directory")
    if(file_name == ""):
        return
    else:

        df_to_merge = pd.read_csv(file_name[0],skiprows=9)
        
        self.df = pd.append({df_to_merge["ID Number"]:self.df["Location"]}, ignore_index=True)
        print(self.df)