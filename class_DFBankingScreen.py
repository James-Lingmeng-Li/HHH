"""Banking

This is where Debtor Finanace processes such as downloading/sending Westpac reports, importing Debtor Receipts and such should go.

Certain scripts are actually completed by calling AutoHotKey (AHK) code, as Aquarius (our 3rd party Debtor Finance application) does not allow direct integration, and AHK serves us better than Python for this
"""

import os, subprocess, datetime, shutil, copy, traceback, re

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

from cryptography.fernet import Fernet
from PyQt5 import QtCore, QtGui, QtWidgets

import HHHconf, HHHfunc

class bankingScreen(HHHconf.HHHWindowWidget):
    signal_addEvent = QtCore.pyqtSignal(object)
    signal_animateWindow = QtCore.pyqtSignal(bool)
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.dict_receiptsStatus = {
            'allocated':'Complete', 
            'unallocated':None, 
            'pending':self.username
        }
        self.addWidgets()

    def addWidgets(self):
        heading_banking = QtWidgets.QLabel('Banking')
        heading_banking.setStyleSheet(HHHconf.design_textMedium)
        heading_banking.setMaximumHeight(20)
        self.vBox_contents.addWidget(heading_banking)

        # Send Segments Box
        self.sendSegmentsBox = QtWidgets.QGroupBox()
        self.sendSegmentsBox.setContentsMargins(5, 5, 5, 5)
        self.sendSegmentsBox.setStyleSheet(HHHconf.design_GroupBox)
        self.vBox_contents.addWidget(self.sendSegmentsBox)

        vBox_sendSegmentsBox = QtWidgets.QVBoxLayout()
        vBox_sendSegmentsBox.setContentsMargins(5, 5, 5, 5)
        vBox_sendSegmentsBox.setSpacing(5)
        vBox_sendSegmentsBox.setAlignment(QtCore.Qt.AlignTop)   
        self.sendSegmentsBox.setLayout(vBox_sendSegmentsBox) 

        hBox_titleSegments =  QtWidgets.QHBoxLayout()
        hBox_titleSegments.setContentsMargins(5, 5, 5, 5)
        hBox_titleSegments.setSpacing(5)
        hBox_titleSegments.setAlignment(QtCore.Qt.AlignTop)   
        vBox_sendSegmentsBox.addLayout(hBox_titleSegments) 

        heading_sendSegments = QtWidgets.QLabel('WPC Segments')
        heading_sendSegments.setStyleSheet(HHHconf.design_textMedium)
        heading_sendSegments.setMaximumHeight(20)
        hBox_titleSegments.addWidget(heading_sendSegments)

        button_sendWPCSegmentReports = QtWidgets.QPushButton('Send WPC Segment Reports (Today Only)', objectName='button_sendWPCSegments', clicked=self.beginSegmentsThread, default=True)  
        button_sendWPCSegmentReports.setStyleSheet(HHHconf.design_smallButtonTransparent)
        button_sendWPCSegmentReports.setFixedHeight(50)
        button_sendWPCSegmentReports.setFocusPolicy(QtCore.Qt.NoFocus)
        vBox_sendSegmentsBox.addWidget(button_sendWPCSegmentReports) 

        # Import Receipts Box
        self.importReceiptsBox = QtWidgets.QGroupBox()
        self.importReceiptsBox.setContentsMargins(5, 5, 5, 5)
        self.importReceiptsBox.setStyleSheet(HHHconf.design_GroupBox)
        self.vBox_contents.addWidget(self.importReceiptsBox)

        vBox_importReceiptsBox = QtWidgets.QVBoxLayout(self.importReceiptsBox)
        vBox_importReceiptsBox.setContentsMargins(5, 5, 5, 5)
        vBox_importReceiptsBox.setSpacing(5)
        vBox_importReceiptsBox.setAlignment(QtCore.Qt.AlignTop)   
        self.importReceiptsBox.setLayout(vBox_importReceiptsBox) 

        hBox_titleReceipts =  QtWidgets.QHBoxLayout()
        hBox_titleReceipts.setContentsMargins(5, 5, 5, 5)
        hBox_titleReceipts.setSpacing(5)
        hBox_titleReceipts.setAlignment(QtCore.Qt.AlignTop)   
        vBox_importReceiptsBox.addLayout(hBox_titleReceipts) 

        heading_importReceipts = QtWidgets.QLabel('Debtor Receipts')
        heading_importReceipts.setStyleSheet(HHHconf.design_textMedium)
        heading_importReceipts.setMaximumHeight(20)
        hBox_titleReceipts.addWidget(heading_importReceipts)

        hBox_titleReceipts.addStretch(1)

        heading_receiptsDate = QtWidgets.QLabel('Date: ', objectName='heading_lastRefreshed')
        heading_receiptsDate.setStyleSheet(HHHconf.design_textSmall.replace(HHHconf.mainFontColor, 'gray').replace('{', '{padding-top: 3px;'))
        heading_receiptsDate.setMaximumHeight(20)
        heading_receiptsDate.setFixedWidth(40)
        hBox_titleReceipts.addWidget(heading_receiptsDate)

        self.calendar_dateRefreshed = HHHconf.calendarLineEditWidget(parent=self.importReceiptsBox, objName='dateRefreshed', fixedWidth=80, calWidth=300, calHeight=300, paddingLeft='3px')
        self.calendar_dateRefreshed.setText(datetime.datetime.now().strftime(HHHconf.dateFormat2))
        hBox_titleReceipts.addWidget(self.calendar_dateRefreshed)

        button_refreshDebtorReceipts = QtWidgets.QPushButton('Refresh', objectName='button_refreshDebtorReceipts', clicked=self.get_debtorReceiptsData, default=True)  
        button_refreshDebtorReceipts.setStyleSheet(HHHconf.design_smallButtonTransparent)
        hBox_titleReceipts.addWidget(button_refreshDebtorReceipts)

        # Debtor Receipts Table
        hBox_debtorReceipts = QtWidgets.QHBoxLayout()
        hBox_debtorReceipts.setContentsMargins(0, 0, 0, 0)
        hBox_debtorReceipts.setSpacing(5)
        hBox_debtorReceipts.setAlignment(QtCore.Qt.AlignTop)   
        vBox_importReceiptsBox.addLayout(hBox_debtorReceipts) 

        self.debtorReceiptsTable = HHHconf.HHHTableWidget(fixedHeight=300, cornerWidget='copy', contextMenu=self.debtorReceiptsTableMenu)
        hBox_debtorReceipts.addWidget(self.debtorReceiptsTable)    

        vBox_debtorReceiptsDetails = QtWidgets.QVBoxLayout()
        vBox_debtorReceiptsDetails.setContentsMargins(0, 0, 0, 0)
        vBox_debtorReceiptsDetails.setSpacing(15)
        vBox_debtorReceiptsDetails.setAlignment(QtCore.Qt.AlignBottom)   
        hBox_debtorReceipts.addLayout(vBox_debtorReceiptsDetails) 

        self.dict_debtorReceiptsDetails = {
            'lastRefreshed':{
                'heading':'Last Refreshed', 
                'format':HHHconf.timeFormat
            }, 
            'totalReceipts':{
                'heading':'Total Receipts', 
                'format':HHHconf.moneyFormat
            }, 
            'allocated':{
                'heading':'Allocated', 
                'format':None
            }, 
            'unallocated':{
                'heading':'Unallocated', 
                'format':None
            }, 
        }
        for obj, options in self.dict_debtorReceiptsDetails.items():
            hBox_current = QtWidgets.QHBoxLayout()
            hBox_current.setContentsMargins(5, 0, 0, 5)
            hBox_current.setSpacing(5)
            vBox_debtorReceiptsDetails.addLayout(hBox_current) 

            currentHeading = QtWidgets.QLabel(options['heading'], objectName='heading_' + str(obj))
            currentHeading.setStyleSheet(HHHconf.design_textSmall)
            currentHeading.setFixedWidth(90)
            hBox_current.addWidget(currentHeading)

            currentValue = QtWidgets.QLabel('', objectName='value_' + str(obj))
            currentValue.setStyleSheet(HHHconf.design_textSmallMonoSpace.replace(HHHconf.mainFontColor, 'gray'))
            hBox_current.addWidget(currentValue)


        button_importDebtorReceipts = QtWidgets.QPushButton('Import Receipts to AQ', objectName='button_importDebtorReceipts', clicked=self.importReceipts, default=True)  
        button_importDebtorReceipts.setStyleSheet(HHHconf.design_smallButtonTransparent)
        button_importDebtorReceipts.setFixedHeight(50)
        button_importDebtorReceipts.setFixedWidth(200)
        vBox_debtorReceiptsDetails.addWidget(button_importDebtorReceipts) 

        button_reverseDebtorReceipts = QtWidgets.QPushButton('Undo Receipts Import', objectName='button_reverseDebtorReceipts', clicked=self.reverseReceipts, default=True)  
        button_reverseDebtorReceipts.setStyleSheet(HHHconf.design_smallButtonTransparent)
        button_reverseDebtorReceipts.setFixedHeight(50)
        button_reverseDebtorReceipts.setFixedWidth(200)
        vBox_debtorReceiptsDetails.addWidget(button_reverseDebtorReceipts) 

        self.get_debtorReceiptsData()        

        # Debtor Receipts Results
        self.resultsBox = QtWidgets.QGroupBox()
        self.resultsBox.setContentsMargins(5, 5, 5, 5)
        self.resultsBox.setStyleSheet(HHHconf.design_GroupBox)
        self.vBox_contents.addWidget(self.resultsBox)

        vBox_resultsBoxStack = QtWidgets.QVBoxLayout()
        vBox_resultsBoxStack.setContentsMargins(5, 5, 5, 5)
        vBox_resultsBoxStack.setSpacing(5)
        vBox_resultsBoxStack.setAlignment(QtCore.Qt.AlignTop)   
        self.resultsBox.setLayout(vBox_resultsBoxStack) 

        vBox_resultsBoxStack.addStretch(1)

    def debtorReceiptsTableMenu(self, pos):
        #TODO: Add context menu actions to exclude or add additional debtor receipt lines
        # menu = QtWidgets.QMenu()
        # menu.setStyleSheet(HHHconf.design_tableMenu)
        
        # action_excludeReceipts = menu.addAction('Exclude selected item')
        # action_addReceipts = menu.addAction('Add item to table')

        # action = menu.exec_(QtGui.QCursor.pos())
        # if action == action_excludeReceipts:
        #     pass
        # elif action == action_addReceipts:
        #     pass
        pass

    def get_debtorReceiptsData(self):
        self.selectedDateString = self.calendar_dateRefreshed.text()
        try:
            self.selectedDate = datetime.datetime.strptime(self.selectedDateString, HHHconf.dateFormat2) 
        except:
            self.clearDebtorReceipts()
            return

        self.df_debtorReceipts = HHHfunc.debtorFinanceEngine.getDebtorReceipts(date=self.selectedDate)
        if len(self.df_debtorReceipts.index) == 0:
            self.clearDebtorReceipts()
            return

        # Populate table
        self.df_debtorReceiptsDisplay = copy.deepcopy(self.df_debtorReceipts)
        self.df_debtorReceiptsDisplay[HHHconf.debtorReceipts_value] = self.df_debtorReceiptsDisplay[HHHconf.debtorReceipts_value].map(HHHconf.moneyFormat.format) # Copy required to display, original is used for metrics calcs

        self.df_debtorReceiptsDisplay.rename({HHHconf.DFClientInfo_AQClientNumber:'AQ Number', HHHconf.debtorReceipts_client:'Client', HHHconf.debtorReceipts_value:'Amount', HHHconf.debtorReceipts_status:'Status'}, axis='columns', inplace=True)
        self.debtorReceiptsTable.setModel(self.df_debtorReceiptsDisplay)
        self.debtorReceiptsTable.sortByColumn(0, QtCore.Qt.AscendingOrder)

        # Get and display metrics
        self.importReceiptsBox.findChild(QtWidgets.QLabel, 'value_lastRefreshed').setText(datetime.datetime.now().strftime(self.dict_debtorReceiptsDetails['lastRefreshed']['format']))

        totalReceiptsCount = len(self.df_debtorReceipts[self.df_debtorReceipts[HHHconf.debtorReceipts_value] > 0].index)
        self.importReceiptsBox.findChild(QtWidgets.QLabel, 'value_totalReceipts').setText(self.dict_debtorReceiptsDetails['totalReceipts']['format'].format(self.df_debtorReceipts[HHHconf.debtorReceipts_value].sum(skipna=True)))

        allocatedCount = len(self.df_debtorReceipts[self.df_debtorReceipts[HHHconf.debtorReceipts_status] == self.dict_receiptsStatus['allocated']].index)
        self.importReceiptsBox.findChild(QtWidgets.QLabel, 'value_allocated').setText(str(allocatedCount) +  ' / ' + str(totalReceiptsCount))

        unallocatedCount = len(self.df_debtorReceipts[(self.df_debtorReceipts[HHHconf.debtorReceipts_status] != self.dict_receiptsStatus['allocated']) & (self.df_debtorReceipts[HHHconf.debtorReceipts_value] > 0)].index)
        self.importReceiptsBox.findChild(QtWidgets.QLabel, 'value_unallocated').setText(str(unallocatedCount) +  ' / ' + str(totalReceiptsCount))

    def importReceipts(self):
        if len(self.df_debtorReceipts.index) == 0:
            return HHHconf.widgetEffects.flashMessage(self.heading_status, duration=3000, message='Receipts for the selected date have not been calculated yet - Please contact accounts', messageType='warning') 

        # import pandas # START TEST (must comment out receiptsWorker too, prevent db status updates during testing)
        # testDict = {
        #     HHHconf.DFClientInfo_AQClientNumber: ['0000023','0000023'],
        #     HHHconf.debtorReceipts_client:['TEST - TT Marquet Pty Ltd','TEST - TT Marquet Pty Ltd'],
        #     HHHconf.debtorReceipts_value:[11111.11,22222.22],
        #     'Status':['None','None']
        # }
        # self.df_debtorReceiptsImport = pandas.DataFrame.from_dict(testDict)
        # print(self.df_debtorReceiptsImport) # END TEST

        self.df_debtorReceiptsImport = copy.deepcopy(self.df_debtorReceipts)
        self.df_debtorReceiptsImport = self.df_debtorReceiptsImport[(self.df_debtorReceiptsImport[HHHconf.debtorReceipts_value] != 0) & (self.df_debtorReceiptsImport[HHHconf.debtorReceipts_status].isnull())]

        self.beginReceiptsThread(self.df_debtorReceiptsImport, self.dict_receiptsStatus['unallocated'])

    def reverseReceipts(self):
        if len(self.df_debtorReceipts.index) == 0:
            return HHHconf.widgetEffects.flashMessage(self.heading_status, duration=3000, message='Receipts for the selected date have not been calculated yet - Please contact accounts', messageType='warning') 
        
        # import pandas # START TEST (must comment out receiptsWorker too, prevent db status updates during testing)
        # testDict = {
        #     HHHconf.DFClientInfo_AQClientNumber: ['0000023','0000023'],
        #     HHHconf.debtorReceipts_client:['TEST - TT Marquet Pty Ltd','TEST - TT Marquet Pty Ltd'],
        #     HHHconf.debtorReceipts_value:[11111.11,22222.22],
        #     'Status':['None','None']
        # }
        # self.df_debtorReceiptsImport = pandas.DataFrame.from_dict(testDict)
        # print(self.df_debtorReceiptsImport) # END TEST

        self.df_debtorReceiptsImport = copy.deepcopy(self.df_debtorReceipts)
        self.df_debtorReceiptsImport = self.df_debtorReceiptsImport[(self.df_debtorReceiptsImport[HHHconf.debtorReceipts_value] != 0) & (self.df_debtorReceiptsImport[HHHconf.debtorReceipts_status] == self.dict_receiptsStatus['allocated'])]

        self.beginReceiptsThread(self.df_debtorReceiptsImport, self.dict_receiptsStatus['allocated'])

    def beginSegmentsThread(self):
        # Password confirm screen
        if not HHHfunc.mainEngine.requestPassword(self, self.username, title='Confirm Action', dialogText='You are about to download and send segment reports. Please enter your password to continue:'):
            return
        # Only one segments worker allowed at a time
        try:
            if self.segmentsThread.isRunning():
                return HHHconf.widgetEffects.flashMessage(self.heading_status, duration=3000, message='Segments Worker already at work. Please wait before trying again.', messageType='warning') 
        except:
            pass
        # Check WPC credentials
        df_appCredentials = HHHfunc.mainEngine.getAppCredentials(username=self.username)
        try:
            dict_WPCCredentials = df_appCredentials[df_appCredentials[HHHconf.userCredentials_application].str.contains('WPC')].to_dict('records')[0]
        except IndexError:
            return HHHconf.widgetEffects.flashMessage(self.heading_status, duration=3000, message='WPC credentials not found. Check your settings', messageType='error')
        currentWPCpassword = HHHfunc.mainEngine.twoWayDecrypt(str(dict_WPCCredentials[HHHconf.userCredentials_password]))

        # Initiate receipts thread (so GUI does not time out)
        self.segmentsThread = QtCore.QThread()
        self.segmentsWorker = HHHfunc.debtorFinanceEngine.segmentsWorker(parent=self, username=self.username, WPCusername=str(dict_WPCCredentials[HHHconf.userCredentials_username]), WPCpassword=currentWPCpassword)
        self.segmentsWorker.moveToThread(self.segmentsThread)
        self.segmentsThread.started.connect(self.segmentsWorker.beginSegments)
        self.segmentsWorker.signal_finished.connect(lambda emittedMessage: HHHconf.widgetEffects.flashMessage(self.heading_status, duration=3000, message='Segment Reports Worker completed. See the status list for details', messageType='success'))
        self.segmentsWorker.signal_finished.connect(self.segmentsThread.quit)
        self.segmentsWorker.signal_finished.connect(lambda x: self.signal_addEvent.emit(x))
        self.segmentsWorker.signal_passwordUpdated.connect(lambda x: HHHfunc.mainEngine.updateAppCredentials(app=dict_WPCCredentials[HHHconf.userCredentials_application], data={HHHconf.userCredentials_password: HHHfunc.mainEngine.twoWayEncrypt(x)}, username=self.username))
        self.segmentsThread.start()
    
    def beginReceiptsThread(self, df_receipts, currentStatus):
        self.action = 'import' if currentStatus == self.dict_receiptsStatus['unallocated'] else 'reverse'
        self.currentStatus = currentStatus
        # Password confirm screen
        if not HHHfunc.mainEngine.requestPassword(self, self.username, title='Confirm Action', dialogText='You are about to <html style="font-weight:bold;">' + self.action + '</html> receipts. Please enter your password to continue:'):
            return
        # Only one receipts worker allowed at a time
        try:
            if self.receiptsThread.isRunning():
                return HHHconf.widgetEffects.flashMessage(self.heading_status, duration=3000, message='Receipts Worker already at work. Please wait before trying again.', messageType='warning') 
        except:
            pass
        # Check Scale Factor - Stop if not 100%
        if HHHfunc.getScaleFactor() != 1:
            return HHHconf.widgetEffects.flashMessage(self.heading_status, duration=3000, message='Windows display scaling needs to be 100% for this process. Check your display settings', messageType='error')

        # Check AQ credentials
        df_appCredentials = HHHfunc.mainEngine.getAppCredentials(username=self.username)
        try:
            dict_AQCredentials = df_appCredentials[df_appCredentials[HHHconf.userCredentials_application].str.contains('AQ')].to_dict('records')[0]
        except IndexError:
            return HHHconf.widgetEffects.flashMessage(self.heading_status, duration=3000, message='AQ credentials not found. Check your settings', messageType='error')

        # Check if AQ installed
        if not os.path.exists(HHHconf.aq_dir):
            return HHHconf.widgetEffects.flashMessage(self.heading_status, duration=3000, message='AQ application not installed. Contact your admin', messageType='error')

        # Initiate receipts thread (so GUI does not time out)
        self.receiptsThread = QtCore.QThread()
        self.receiptsWorker = HHHfunc.debtorFinanceEngine.receiptsWorker(df_receipts=df_receipts, currentStatus=self.currentStatus, action=self.action, parent=self, username=dict_AQCredentials[HHHconf.userCredentials_username],password=HHHfunc.mainEngine.twoWayDecrypt(str(dict_AQCredentials[HHHconf.userCredentials_password])))
        self.receiptsWorker.moveToThread(self.receiptsThread)
        self.receiptsThread.started.connect(self.receiptsWorker.beginReceipts)
        self.receiptsWorker.signal_started.connect(lambda: self.animateWindowReceipts(True, self.currentStatus, self.action))
        self.receiptsWorker.signal_finished.connect(lambda: self.animateWindowReceipts(False, self.currentStatus, self.action)) 
        self.receiptsWorker.signal_finished.connect(self.receiptsThread.quit)
        self.receiptsWorker.signal_finished.connect(lambda x: self.signal_addEvent.emit(x))
        self.receiptsThread.start()
        return

    def animateWindowReceipts(self, state, currentstatus, action):
        if state is False:
            # Return pending status back to original status if the import was not successfully completed (e.g. if user exits early)
            HHHfunc.debtorFinanceEngine.updateDebtorReceiptStatus(oldStatus=self.dict_receiptsStatus['pending'], newStatus=self.currentStatus, date=self.selectedDate)
            self.get_debtorReceiptsData()
        # Return window state back to normal            
        HHHconf.widgetEffects.translucifyWindow(self, state, self.geometry())
        self.signal_animateWindow.emit(state)
        HHHconf.widgetEffects.flashMessage(self.heading_status, duration=3000, message='Debtor receipts <html style="font-weight:bold;">' + action + '</html> complete.', messageType='success')
    
    def clearDebtorReceipts(self):
        if self.debtorReceiptsTable.model():
            self.debtorReceiptsTable.model().deleteLater()
        for child in self.importReceiptsBox.findChildren(QtWidgets.QLabel):
            if 'value_' in child.objectName(): 
                child.setText('')
        