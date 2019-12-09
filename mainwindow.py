#-*- coding: utf-8 -*-
import sys
reload(sys)
sys.setdefaultencoding('utf-8')
import sip
sip.setapi('QString', 2)
sip.setapi('QVariant', 2)
from PyQt4.QtCore import *
from PyQt4.QtGui import *
import crwaller


class Screenshot(QWidget):
    MESSAGE = 'Job Done!'
    def __init__(self):
        super(Screenshot, self).__init__()
        self.filename = 'cosmicgirl.xlsx'
        browseButton = self.createButton("&Browse...", self.browse)
        self.screenshotLabel = QLabel()
        self.screenshotLabel.setSizePolicy(QSizePolicy.Expanding,
                QSizePolicy.Expanding)
        self.screenshotLabel.setAlignment(Qt.AlignCenter)
        self.screenshotLabel.setMinimumSize(240, 160)
        self.directoryComboBox = self.createComboBox(QDir.currentPath())
        self.createOptionsGroupBox()
        self.createpsoptionbox()
        self.createButtonsLayout()
        self.processButton.clicked.connect(self.process)
        mainLayout = QVBoxLayout()
        mainLayout.addWidget(self.directoryComboBox)
        mainLayout.addWidget(browseButton)
        mainLayout.addWidget(self.optionsGroupBox)
        mainLayout.addWidget(self.optionsGroupBox2)
        mainLayout.addLayout(self.buttonsLayout)

        self.setLayout(mainLayout)


        self.setWindowTitle("WSJN")
        self.resize(300, 200)

    def browse(self):
        directory = QFileDialog.getExistingDirectory(self, "Find Files",
                QDir.currentPath())

        if directory:
            if self.directoryComboBox.findText(directory) == -1:
                self.directoryComboBox.addItem(directory)

            self.directoryComboBox.setCurrentIndex(self.directoryComboBox.findText(directory))

    def createComboBox(self, text=""):
        comboBox = QComboBox()
        comboBox.setEditable(True)
        comboBox.addItem(text)
        comboBox.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        return comboBox

    def createOptionsGroupBox(self):
        self.optionsGroupBox = QGroupBox("Options")
        self.datelineEdit = QLineEdit()
        self.datelineEdit2 = QLineEdit()
        self.delaySpinBoxLabel2 = QLabel("Start Date : ")
        self.delaySpinBoxLabel = QLabel("End Date : ")
        self.galllineEdit = QLineEdit("wjsnvlive")
        self.gallSpinBoxLabel = QLabel("Gallery : ")
        self.minorBoxLabel = QLabel("Minor gall : ")
        self.checkminor = QCheckBox()
        self.checkminor.setChecked(1)

        self.hideThisWindowCheckBox = QCheckBox("Hide This Window")

        optionsGroupBoxLayout = QGridLayout()
        optionsGroupBoxLayout.addWidget(self.gallSpinBoxLabel,1,0,1,2)
        optionsGroupBoxLayout.addWidget(self.galllineEdit,1,1,1,2)
        optionsGroupBoxLayout.addWidget(self.delaySpinBoxLabel2,2,0,1,2)
        optionsGroupBoxLayout.addWidget(self.datelineEdit2,2,1,1,2)
        optionsGroupBoxLayout.addWidget(self.delaySpinBoxLabel,3,0,1,2)
        optionsGroupBoxLayout.addWidget(self.datelineEdit,3,1,1,2)
        optionsGroupBoxLayout.addWidget(self.minorBoxLabel, 4, 0, 1, 2)
        optionsGroupBoxLayout.addWidget(self.checkminor, 4, 1, 1, 2)

        self.optionsGroupBox.setLayout(optionsGroupBoxLayout)


    def createpsoptionbox(self):
        self.optionsGroupBox2 = QGroupBox("Filter")
        self.prefixLabel = QLabel("Pre : ")
        self.preEdit = QLineEdit()
        self.suffixLabel = QLabel("Suf : ")
        self.sufEdit = QLineEdit()

        optionspsLayout = QGridLayout()
        optionspsLayout.addWidget(self.prefixLabel, 1,0,1,2)
        optionspsLayout.addWidget(self.preEdit, 1, 1, 1, 2)
        optionspsLayout.addWidget(self.suffixLabel, 2,0,1,2)
        optionspsLayout.addWidget(self.sufEdit, 2, 1, 1, 2)
        self.optionsGroupBox2.setLayout(optionspsLayout)


    def createButtonsLayout(self):
        self.processButton = QPushButton("Process")
        self.quitScreenshotButton = self.createButton("Quit", self.close)

        self.buttonsLayout = QHBoxLayout()
        self.buttonsLayout.addStretch()
        self.buttonsLayout.addWidget(self.processButton)
        self.buttonsLayout.addWidget(self.quitScreenshotButton)

    def informationMessage(self):
        reply = QMessageBox.information(self,
                "Finished", Screenshot.MESSAGE)
        if reply == QMessageBox.Ok:
            self.informationLabel.setText("OK")
        else:
            self.informationLabel.setText("Escape")


    def process(self):
        parser = crwaller.DcInsideParser("chromedriver.exe", self.galllineEdit.text())
        parser.set_date(self.datelineEdit.text(), self.datelineEdit2.text())
        parser.set_filter(self.preEdit.text(), self.sufEdit.text())
        i = 1
        minor = False
        while (True):
            parser.set_page_no(i)
            if self.checkminor.isChecked():
                minor = True
            parser.load_document(minor)
            parser.load_posts()
            print('page no : ' + str(i))
            print('page flag :' + str(parser.pageflag))
            if parser.pageflag == 0:
                break
            i += 1
        parser.creatExcelfile(self.filename, str(self.directoryComboBox.currentText()))
        print('last page was ' + str(i))
        self.informationMessage()

    def createButton(self, text, member):
        button = QPushButton(text)
        button.clicked.connect(member)
        return button

if __name__ == '__main__':
    import sys
    app = QApplication(sys.argv)
    screenshot = Screenshot()
    screenshot.show()
    sys.exit(app.exec_())
