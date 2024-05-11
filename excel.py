import sys
import numpy as np
import pyqtgraph as pg
import pandas as pd
from scipy import signal

from PyQt6.QtWidgets import QApplication, QMainWindow, QFileDialog, QListWidgetItem, QComboBox
from PyQt6.QtGui import QPainter, QColor, QFont
from PyQt6.QtCore import QThread, pyqtSignal, pyqtSignal, QObject
from app_ui import Ui_MainWindow


class ExcelImporter:
    def __init__(self, app):
        self.app = app
        self.excelFilepath = None
        self.excelDataFrame = None

    def importExcelFile(self):
        dialog = QFileDialog(self.app)
        dialog.setDirectory(".")
        dialog.setNameFilter("Excel / CSV files (*.xlsx *csv)")
        dialog.setViewMode(QFileDialog.ViewMode.List)
        if dialog.exec() == 1:
            filenames = dialog.selectedFiles()
            self.excelFilepath = str(filenames[0])
            self.parseExcelFile()

    def parseExcelFile(self):
        if self.excelFilepath is not None:
            header = [None, 0][int(self.app.ui.cbExcelHasColumnNames.isChecked())]
            index_col = [None, 0][int(self.app.ui.cbExcelHasIndexColumn.isChecked())]
            self.excelDataFrame = pd.read_excel(self.excelFilepath, header=header, index_col=index_col)
            self.resetExcelTable()

    def resetExcelTable(self):
        self.app.ui.tableExcelDataSelection.setRowCount(0)

    def addRowToExcelTable(self):
        if self.excelDataFrame is not None:
            row_count = self.app.ui.tableExcelDataSelection.rowCount()
            self.app.ui.tableExcelDataSelection.setRowCount(row_count + 1)

            row_count = self.app.ui.tableExcelDataSelection.rowCount()
            columns = self.excelDataFrame.columns
            cmb_x = QComboBox()
            cmb_y = QComboBox()
            for i in range(len(columns)):
                cmb_x.addItem(str(columns[i]), userData=i)
                cmb_y.addItem(str(columns[i]), userData=i)
            self.app.ui.tableExcelDataSelection.setCellWidget(row_count - 1, 0, cmb_x)
            self.app.ui.tableExcelDataSelection.setCellWidget(row_count - 1, 1, cmb_y)

    def popRowFromExcelTable(self):
        row_count = self.app.ui.tableExcelDataSelection.rowCount()
        if row_count >= 1:
            self.app.ui.tableExcelDataSelection.setRowCount(row_count - 1)

    def updateDataFromExcel(self):
        if self.excelDataFrame is not None:
            data = []
            labels = []
            columns = self.excelDataFrame.columns
            for row in range(self.app.ui.tableExcelDataSelection.rowCount()):
                x_label = columns[self.app.ui.tableExcelDataSelection.cellWidget(row, 0).currentData()]
                y_label = columns[self.app.ui.tableExcelDataSelection.cellWidget(row, 1).currentData()]
                data.append([np.array(self.excelDataFrame[x_label]), np.array(self.excelDataFrame[y_label])])
                labels.append(y_label)
            # self.signals.dataChanged.emit()
            self.app.data = data
            self.app.labels = labels
