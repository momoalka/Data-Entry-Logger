import os
import sys

from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QComboBox, QLineEdit, QPushButton, QHBoxLayout, QVBoxLayout, QDateEdit, \
    QAbstractSpinBox, QCalendarWidget, QDateTimeEdit

from PyQt5.QtCore import Qt, QDate, QTime
from PyQt5.QtGui import QIcon

import win32com.client as win32
import win32api
import win32com

class SubmittalsRFisLog:
    def __init__(self, parent=None):
        self.parent = parent
        self.xlApp = win32.Dispatch('Excel.Application')
        self.xlApp.Visible = True

        # Excel Workbook Reference
        self.wb = self.xlApp.Workbooks.Open(os.path.join(os.getcwd(), 'BF-54-2019 Submittal Log.xlsx'))

        # Excel Worksheet object Reference
        self.ws = self.wb.Worksheets('Submittals')
        self.parent.status.setText('Submittal & RFI Log Connected')

    def addEntry(self, record: list=None):

        try:
            rowIndx = self.ws.Cells(self.ws.Rows.Count, "A").End(-4162).Row
            '''
            #rowIndx = self.ws.Range('A1:A' + str(lastrow)).Find('*', self.ws.Cells(lastrow, "A"), -4163, 2, 1, 2).Row
            '''
            rowIndx += 1
            self.ws.Range(
                self.ws.Cells(rowIndx, "A"),
                self.ws.Cells(rowIndx, "K"),
            ).Value = record
            self.wb.Save()

        except Exception as e:
            self.parent.status.setText(win32api.FormatMessage(e.hrsult))
'''
    def closeFile(self, value: bool=False):

        if value == True:
            print('doing these')
            self.wb.Save()
            self.wb.Close(SaveChanges = True)
            self.xlApp.Quit()
'''

class TimeEntryField(QDateTimeEdit):
    def __init__(self):
        super().__init__()
        self.setDisplayFormat('HH:mm AP')

    def stepBy(self, steps):
        if self.currentSection() == QDateTimeEdit.MinuteSection:
            self.setTime(self.time().addSecs(60*15*steps))
            return
        super().stepBy(steps)

class DataEntryApp(QWidget):

    def __init__(self):
        super().__init__()
        self.setWindowTitle('BFB-15-2019: Main Cable Dehumidification')
        self.setWindowIcon(QIcon('TTLogo.ico'))


        self.setMinimumWidth(500)

        self.layout = QVBoxLayout()
        self.layout.setSpacing(20)
        self.setLayout(self.layout)

        self.initUI()

        self.reset_fields()
        self.SubsRFIs = SubmittalsRFisLog(self)
        #self.close_File(self)


    def initUI(self):

        subLayouts = {}

        subLayouts['DateTime'] = QHBoxLayout()
        self.layout.addLayout(subLayouts['DateTime'])

        labelDate = QLabel('Date: ')
        self.lineEditDate = QDateEdit()
        self.lineEditDate.setCalendarPopup(True)

        labelTime = QLabel('Time: ')
        self.lineEditTime = TimeEntryField()

        subLayouts['DateTime'].addWidget(labelDate, 0, alignment=Qt.AlignRight)
        subLayouts['DateTime'].addWidget(self.lineEditDate, 5)
        subLayouts['DateTime'].addWidget(labelTime, 0, alignment=Qt.AlignRight)
        subLayouts['DateTime'].addWidget(self.lineEditTime, 5)


        # Row 2

        subLayouts[2] =  QHBoxLayout()
        self.layout.addLayout(subLayouts[2])

        self.comboReviewerClassification = QComboBox()
        self.comboReviewerClassification.setMaximumWidth(int(self.rect().width()*0.2))
        self.comboReviewerClassification.addItems(('TT', 'COWI', 'WCA'))
        subLayouts[2].addWidget(QLabel('Reviewer: '), 0, alignment = Qt.AlignRight)
        subLayouts[2].addWidget(self.comboReviewerClassification, 5)

        self.comboTypeClassification = QComboBox()
        self.comboTypeClassification.setMaximumWidth(int(self.rect().width()*0.2))
        self.comboTypeClassification.addItems(('Submittal', 'RFI'))
        subLayouts[2].addWidget(QLabel('Type: '), 0, alignment = Qt.AlignRight)
        subLayouts[2].addWidget(self.comboTypeClassification, 5)



        '''
        self.lineEditBusinessDetails = QLineEdit()
        subLayouts[2].addWidget(QLabel('Business Details: '), 2, alignment = Qt.AlignRight)
        subLayouts[2].addWidget(self.lineEditBusinessDetails, 8)
        '''

        subLayouts[3] = QHBoxLayout()
        subLayouts[4] = QHBoxLayout()
        subLayouts[5] = QHBoxLayout()

        self.layout.addLayout(subLayouts[3])
        self.layout.addLayout(subLayouts[4])
        self.layout.addLayout(subLayouts[5])

        self.lineSubmittalName = QLineEdit()
        subLayouts[3].addWidget(QLabel('Submittal Name: '), 2, alignment=Qt.AlignRight)
        subLayouts[3].addWidget(self.lineSubmittalName, 8)

        self.lineEditSubject = QLineEdit()
        subLayouts[4].addWidget(QLabel('Subject: '), 2, alignment=Qt.AlignRight)
        subLayouts[4].addWidget(self.lineEditSubject, 8)

        self.lineEditRevisionNo = QLineEdit()
        subLayouts[5].addWidget(QLabel('Revision Number: '), 2, alignment=Qt.AlignRight)
        subLayouts[5].addWidget(self.lineEditRevisionNo, 8, alignment=Qt.AlignLeft)

        subLayouts['buttons'] = QHBoxLayout()
        subLayouts['buttons'].addStretch()
        self.layout.addLayout(subLayouts['buttons'])

        buttonEnter = QPushButton('&Enter', clicked=self.add_entry) #TODO
        subLayouts['buttons'].addWidget(buttonEnter)

        buttonReset = QPushButton('&Reset', clicked=self.reset_fields)
        subLayouts['buttons'].addWidget(buttonReset)

        buttonClose = QPushButton('&Close')
        buttonClose.clicked.connect(app.quit)
        subLayouts['buttons'].addWidget(buttonClose)

        self.status = QLabel()
        self.status.setStyleSheet('''
            font-size: 23 px;
            color: #d4451d;
        ''')

        self.layout.addWidget(self.status, alignment=Qt.AlignLeft)

    def add_entry(self):

        record = [
            self.lineSubmittalName.text(),
            self.lineEditSubject.text(),
            self.lineEditRevisionNo.text(),
            self.lineEditDate.text(),
            self.lineEditDate.text(),
            self.lineEditTime.text(),
            self.comboReviewerClassification.currentText(),
            self.comboTypeClassification.currentText(),

        ]

        self.SubsRFIs.addEntry(record)
        self.reset_fields()
        self.status.setText('Entry Added')


    def reset_fields(self):

        self.lineEditDate.setDate(QDate.currentDate())
        self.lineEditTime.setTime(QTime.currentTime())
        self.comboReviewerClassification.setCurrentIndex(0)
        self.comboTypeClassification.setCurrentIndex(0)

        #self.lineEditBusinessDetails.clear()
        self.lineSubmittalName.clear()
        self.lineEditSubject.clear()
        self.lineEditRevisionNo.clear()
        self.status.setText('All Fields Reset')
'''
    def close_File(self, value: bool=False):

        self.SubsRFIs.closeFile(value)
'''

if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle('''
        QWidget {
            font-size: 30 px;
            
        }
        QComboBox {
            width: 500 px;
            
        }
        QPushButton {
            width: 200 px
            height: 45 px
        }
        
    ''')

    myApp = DataEntryApp()
    myApp.show()

    try:
        sys.exit(app.exec_())
    except SystemExit:
        print('Closing Window...')
