import sys
import numpy as np
import pyqtgraph as pg
import pandas as pd
from scipy import signal
from spectrum import pcovar
import matplotlib.pyplot as plt

from PyQt6.QtWidgets import QApplication, QMainWindow, QFileDialog, QListWidgetItem, QComboBox
from PyQt6.QtGui import QPainter, QColor, QFont
from PyQt6.QtCore import QThread, pyqtSignal, pyqtSignal, QObject
from PyQt6 import QtCore
from app_ui import Ui_MainWindow

from excel import ExcelImporter


class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        self.excelImporter = ExcelImporter(self)
        self.data: list[list[np.ndarray]] | None = None
        self.labels = []
        self.regionOfInterest = pg.LinearRegionItem([0, 1])
        self.curvesTime = []
        self.curvesFreq = []
        self.colors = ["#0072BD", "#D95319", "#EDB120", "#7E2F8E", "#77AC30", "#77AC30", "#A2142F"]

        self.initDataTab()
        self.initAnalysisTab()

    def initDataTab(self):
        self.ui.tableExcelDataSelection.setColumnCount(2)
        self.ui.tableExcelDataSelection.setHorizontalHeaderLabels(["X", "Y"])

        self.ui.pbImportFromExcel.clicked.connect(self.excelImporter.importExcelFile)
        self.ui.cbExcelHasIndexColumn.stateChanged.connect(self.excelImporter.parseExcelFile)
        self.ui.cbExcelHasColumnNames.stateChanged.connect(self.excelImporter.parseExcelFile)
        self.ui.tbAddRow.clicked.connect(self.excelImporter.addRowToExcelTable)
        self.ui.tbPopRow.clicked.connect(self.excelImporter.popRowFromExcelTable)
        self.ui.pbUpdateDataFromExcel.clicked.connect(self.excelImporter.updateDataFromExcel)

    def initAnalysisTab(self):
        self.ui.plotTimeDomain.showGrid(True, True)
        self.ui.plotFreqDomain.showGrid(True, True)

        self.ui.plotTimeDomain.getAxis('bottom').setStyle(tickFont=QFont("Calibri", 12))
        self.ui.plotTimeDomain.getAxis('left').setStyle(tickFont=QFont("Calibri", 12))
        self.ui.plotFreqDomain.getAxis('bottom').setStyle(tickFont=QFont("Calibri", 12))
        self.ui.plotFreqDomain.getAxis('left').setStyle(tickFont=QFont("Calibri", 12))

        self.ui.plotTimeDomain.addLegend(labelTextSize="12pt", offset=(-20, 20))
        self.ui.plotFreqDomain.addLegend(labelTextSize="12pt", offset=(-20, 20))

        self.ui.plotTimeDomain.addItem(self.regionOfInterest)
        self.regionOfInterest.setZValue(10)
        self.regionOfInterest.hide()
        self.regionOfInterest.sigRegionChangeFinished.connect(self.updateRegionOfInterest)
        self.ui.gbRegionOfInterest.clicked.connect(self.updateRegionOfInterest)

        self.ui.pbUpdateDataFromExcel.clicked.connect(self.updateTimePlot)
        self.ui.pbUpdateDataFromExcel.clicked.connect(self.updateFreqPlot)

        self.ui.gbRegionOfInterest.clicked.connect(self.updateFreqPlot)
        self.regionOfInterest.sigRegionChangeFinished.connect(self.updateFreqPlot)

        self.ui.cbYLogScale.clicked.connect(self.updateFreqPlot)

        self.initWelch()
        self.initPcovar()

    def initWelch(self):
        # windows
        self.ui.cmbWelchWindow.addItem("Прямоугольное окно", "boxcar")
        self.ui.cmbWelchWindow.addItem("Окно Бартлетта", "bartlett")
        self.ui.cmbWelchWindow.addItem("Окно Хэмминга", "hamming")
        self.ui.cmbWelchWindow.addItem("Окно Тейлора", "taylor")
        self.ui.cmbWelchWindow.setCurrentIndex(0)
        self.ui.cmbWelchWindow.currentIndexChanged.connect(self.updateFreqPlot)
        # nperseg
        for n in [128, 256, 512, 1024, 2048, 4096]:
            self.ui.cmbWelchNperseg.addItem(str(n), n)
        self.ui.cmbWelchNperseg.setCurrentIndex(4)
        self.ui.cmbWelchNperseg.currentIndexChanged.connect(self.updateFreqPlot)

        self.ui.gbWelch.clicked.connect(self.updateFreqPlot)
        self.ui.gbWelch.setChecked(False)

    def initPcovar(self):
        self.ui.gbPcovar.clicked.connect(self.updateFreqPlot)
        # nperseg
        for n in [128, 256, 512, 1024, 2048, 4096]:
            self.ui.cmbPcovarNperseg.addItem(str(n), n)
        self.ui.cmbPcovarNperseg.setCurrentIndex(4)
        self.ui.cmbPcovarNperseg.currentIndexChanged.connect(self.updateFreqPlot)
        self.ui.sbPcovarOrder.valueChanged.connect(self.updateFreqPlot)

    def updateTimePlot(self):
        if self.data is not None:
            for curve in self.curvesTime:
                self.ui.plotTimeDomain.removeItem(curve)
            self.curvesTime = []
            for i in range(len(self.data)):
                pen = pg.mkPen(self.colors[i], width=1, style=QtCore.Qt.PenStyle.SolidLine)
                name = self.labels[i]
                self.curvesTime.append(pg.PlotCurveItem(self.data[i][0], self.data[i][1], pen=pen, name=name))
                self.ui.plotTimeDomain.addItem(self.curvesTime[-1])
            x_min = self.data[0][0][0]
            x_max = self.data[0][0][-1]
            self.regionOfInterest.setRegion((x_min, x_max))
            # self.ui.plotTimeDomain.enableAutoRange()

    def updateRegionOfInterest(self):
        if self.ui.gbRegionOfInterest.isChecked():
            self.regionOfInterest.show()
            x1 = float(self.regionOfInterest.getRegion()[0])
            x2 = float(self.regionOfInterest.getRegion()[1])
            dx = float(self.regionOfInterest.getRegion()[1] - self.regionOfInterest.getRegion()[0])
            self.ui.lblROI_x1.setText(f"x1 = {x1:<.6}")
            self.ui.lblROI_x2.setText(f"x2 = {x2:<.6}")
            self.ui.lblROI_dx.setText(f"dx = {dx:<.6}")
        else:
            self.regionOfInterest.hide()

    def updateFreqPlot(self):
        if self.data is not None:
            # self.ui.plotFreqDomain.setLogMode(y=self.ui.cbYLogScale.isChecked())
            self.ui.plotFreqDomain.setLogMode(False, self.ui.cbYLogScale.isChecked())
            for curve in self.curvesFreq:
                self.ui.plotFreqDomain.removeItem(curve)
            self.curvesFreq = []
            if self.ui.gbWelch.isChecked():
                self.plotWelchPeriodogram()
            if self.ui.gbPcovar.isChecked():
                self.plotPcovar()
            self.ui.plotFreqDomain.enableAutoRange()


    def plotWelchPeriodogram(self):
        for i in range(len(self.data)):
            fs = np.median(1 / (self.data[i][0][1:] - self.data[i][0][:-1]))
            if self.regionOfInterest.isVisible():
                indexes = np.where((self.data[i][0] >= self.regionOfInterest.getRegion()[0]) & (
                        self.data[i][0] <= self.regionOfInterest.getRegion()[1]))
                x = self.data[i][0][indexes]
                y = self.data[i][1][indexes]
            else:
                x = self.data[i][0]
                y = self.data[i][1]
            nperseg = min(len(x), self.ui.cmbWelchNperseg.currentData())
            window = self.ui.cmbWelchWindow.currentData()
            f, Pxx_den = signal.welch(y, fs, nperseg=nperseg, window=window)
            f= f[4:]
            Pxx_den = Pxx_den[4:]
            pen = pg.mkPen(self.colors[i], width=1, style=QtCore.Qt.PenStyle.DashLine)
            name = f"{self.labels[i]} Spec. Den. (Welch)"
            curve = pg.PlotDataItem(f, Pxx_den/max(Pxx_den), pen=pen, name=name)
            self.curvesFreq.append(curve)
            self.ui.plotFreqDomain.addItem(self.curvesFreq[-1])

    def plotPcovar(self):
        for i in range(len(self.data)):
            if self.regionOfInterest.isVisible():
                indexes = np.where((self.data[i][0] >= self.regionOfInterest.getRegion()[0]) & (
                        self.data[i][0] <= self.regionOfInterest.getRegion()[1]))
                x = self.data[i][0][indexes]
                y = self.data[i][1][indexes]
            else:
                x = self.data[i][0]
                y = self.data[i][1]
            # nw = 48  # order of an autoregressive prediction model for the signal, used in estimating the PSD.
            # nfft = 256  # NFFT (int) – total length of the final data sets (padded with zero if needed
            nfft = min(len(x), self.ui.cmbPcovarNperseg.currentData())
            order = self.ui.sbPcovarOrder.value()
            fs = np.median(1 / (self.data[i][0][1:] - self.data[i][0][:-1]))
            p = pcovar(y, order, nfft, fs)
            f = np.array(p.frequencies())
            psd = p.psd[4:]
            f = f[4:]
            pen = pg.mkPen(self.colors[i], width=1, style=QtCore.Qt.PenStyle.SolidLine)
            name = f"{self.labels[i]} Spec. Den. (pcovar)"
            curve = pg.PlotDataItem(f, psd/np.max(psd), pen=pen, name=name)
            self.curvesFreq.append(curve)
            self.ui.plotFreqDomain.addItem(self.curvesFreq[-1])

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
