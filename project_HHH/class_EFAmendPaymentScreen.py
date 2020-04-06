
""" EF - Amend Payment Screen

This screen allows users to amend certain scheduled payments

"""
import math, pandas, datetime, openpyxl

from PyQt5 import QtCore, QtGui, QtWidgets

import HHHconf, HHHfunc

class EFAmendPaymentScreen(HHHconf.HHHWindowWidget):

    signal_paymentUpdated = QtCore.pyqtSignal(object)

    def __init__(self, selectedAgreement, selectedPayment, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.selectedAgreement = str(selectedAgreement)
        self.selectedPayment = selectedPayment
        self.addWidgets()
        self.get_EFPaymentDetails()

    def showEvent(self, event):
        try:
            self.childWindow.setFocus()
        except:
            pass

    def closeEvent(self, event):
        try:
            self.childWindow.close()
        except:
            pass

    def moveEvent(self, event):
        try:
            self.childWindow.move(self.geometry().x(), self.geometry().y())
        except:
            pass

    def addWidgets(self): 
        # Payment Details
        hBox_paymentDetails = QtWidgets.QHBoxLayout()
        hBox_paymentDetails.setContentsMargins(0, 0, 0, 0)
        hBox_paymentDetails.setSpacing(10)
        hBox_paymentDetails.setAlignment(QtCore.Qt.AlignTop)   
        self.vBox_contents.addLayout(hBox_paymentDetails)   

        vBox_paymentDetails1 = QtWidgets.QVBoxLayout()
        vBox_paymentDetails1.setContentsMargins(0, 0, 0, 0)
        vBox_paymentDetails1.setSpacing(2)
        vBox_paymentDetails1.setAlignment(QtCore.Qt.AlignTop)   
        hBox_paymentDetails.addLayout(vBox_paymentDetails1)   

        paymentDetailsWidgetList1 = [ #[object_name, heading text, readOnly]
            [HHHconf.agreement_number, 'Agreement Number', True],  
            [HHHconf.client_name, 'Client Name', True], 
            [HHHconf.payment_number, 'Payment Number', True], 
            [HHHconf.repayment_date, 'Original Repayment Date', True], 
            ['outstandingSubordinates', 'Outstanding Subordinates', True], 
        ]
        for detail in paymentDetailsWidgetList1:
            headingName = 'Heading_EFPaymentView_' + detail[0]
            textboxName = 'Textbox_EFPaymentView_' + detail[0]
            self.headingName = QtWidgets.QLabel(detail[1], objectName=headingName)
            self.headingName.setStyleSheet(HHHconf.design_textSmall)
            vBox_paymentDetails1.addWidget(self.headingName)
            self.textboxName = QtWidgets.QLineEdit('', objectName=textboxName, readOnly=detail[2])
            self.textboxName.resize(200, 20)
            self.textboxName.setStyleSheet(HHHconf.design_editBoxTwo)
            vBox_paymentDetails1.addWidget(self.textboxName)    

        vBox_paymentDetails2 = QtWidgets.QVBoxLayout()
        vBox_paymentDetails2.setContentsMargins(0, 0, 0, 0)
        vBox_paymentDetails2.setSpacing(2)
        vBox_paymentDetails2.setAlignment(QtCore.Qt.AlignTop)   
        hBox_paymentDetails.addLayout(vBox_paymentDetails2)

        paymentDetailsWidgetList2 = [ #[object_name, heading text, readOnly]
            [HHHconf.opening_balance, 'Opening Balance', True],  
            [HHHconf.interest_component, 'Interest', True], 
            [HHHconf.principal_component, 'Principal', True], 
            [HHHconf.periodic_repayment, 'Repayment Amount', True], 
            [HHHconf.closing_balance, 'Closing Balance', True], 
        ]
        for detail in paymentDetailsWidgetList2:
            headingName = 'Heading_EFPaymentView_' + detail[0]
            textboxName = 'Textbox_EFPaymentView_' + detail[0]
            self.headingName = QtWidgets.QLabel(detail[1], objectName=headingName)
            self.headingName.setStyleSheet(HHHconf.design_textSmall)
            vBox_paymentDetails2.addWidget(self.headingName)
            self.textboxName = QtWidgets.QLineEdit('', objectName=textboxName, readOnly=detail[2])
            self.textboxName.resize(200, 20)
            self.textboxName.setStyleSheet(HHHconf.design_editBoxTwo)
            vBox_paymentDetails2.addWidget(self.textboxName)

        # Subordinate Account for fees
        vBox_paymentSubordinate = QtWidgets.QVBoxLayout()
        vBox_paymentSubordinate.setContentsMargins(0, 0, 0, 0)
        vBox_paymentSubordinate.setSpacing(1)
        vBox_paymentSubordinate.setAlignment(QtCore.Qt.AlignTop)   
        self.vBox_contents.addLayout(vBox_paymentSubordinate)   

        self.EFPaymentSubordinateTable = HHHconf.HHHTableWidget(cornerWidget='copy', doubleClicked=self.EFPaymentSubordinateTable_doubleClicked)
        vBox_paymentSubordinate.addWidget(self.EFPaymentSubordinateTable)

        self.button_addSubordinate = QtWidgets.QPushButton('Add New Subordinate', clicked=self.addSubordinate, default=True)
        self.button_addSubordinate.setFixedHeight(20)
        self.button_addSubordinate.setStyleSheet(HHHconf.design_smallButtonTransparent)
        vBox_paymentSubordinate.addWidget(self.button_addSubordinate)

        # Payment Notes
        hBox_paymentNotes = QtWidgets.QHBoxLayout()
        hBox_paymentNotes.setContentsMargins(0, 0, 0, 0)
        hBox_paymentNotes.setSpacing(10)
        hBox_paymentNotes.setAlignment(QtCore.Qt.AlignTop)   
        self.vBox_contents.addLayout(hBox_paymentNotes)   
        self.textEdit_paymentNotes = HHHconf.HHHTextEditWidget(self, placeholderText=HHHconf.widgetPlaceHolderText_notes)
        hBox_paymentNotes.addWidget(self.textEdit_paymentNotes)
        self.button_update.clicked.connect(self.update_paymentDetails)
        
    def get_EFPaymentDetails(self):
         #Get Data from db
        cnxn = HHHfunc.mainEngine.establishSqlConnection()
        cursor = cnxn.cursor()

        Job_getEFPaymentDetails = ('select * from ' + HHHconf.db_payoutSchedule_name + ' join ' + HHHconf.db_agreementInfo_name + ' on ' + HHHconf.db_agreementInfo_name + '.' + HHHconf.dict_toSQL_HHH_EF_Agreement_Info[HHHconf.agreement_number] + ' = ' + HHHconf.db_payoutSchedule_name + '.' + HHHconf.dict_toSQL_HHH_EF_Schedule[HHHconf.agreement_number] + ' where ' + HHHconf.db_payoutSchedule_name + '.' + HHHconf.dict_toSQL_HHH_EF_Schedule[HHHconf.payment_number] + ' = ?')
        cursor.execute(Job_getEFPaymentDetails, self.selectedPayment)
        EFPaymentDetails = list(cursor.fetchone())
        EFPaymentDetailsColumns = [column[0] for column in cursor.description]
        for i, columnname in enumerate(EFPaymentDetailsColumns): # rename columns based on dictionary
            if columnname in HHHconf.dict_fromSQL_HHH_EF_Agreement_Info:
                EFPaymentDetailsColumns[i] = HHHconf.dict_fromSQL_HHH_EF_Agreement_Info[columnname]
        for i, columnname in enumerate(EFPaymentDetailsColumns): # rename columns based on dictionary
            if columnname in HHHconf.dict_fromSQL_HHH_EF_Schedule:
                EFPaymentDetailsColumns[i] = HHHconf.dict_fromSQL_HHH_EF_Schedule[columnname]       
        dict_EFPaymentDetails = dict(zip(EFPaymentDetailsColumns, EFPaymentDetails))

        # Display data
        for key in dict_EFPaymentDetails:
            if self.findChild(QtWidgets.QLineEdit, 'Textbox_EFPaymentView_' + str(key)) is None:
                continue
            if key in [HHHconf.opening_balance, HHHconf.interest_component, HHHconf.principal_component, HHHconf.periodic_repayment, HHHconf.closing_balance]:
                if key == HHHconf.periodic_repayment:
                    amount1 = HHHconf.moneyFormat.format(math.ceil(dict_EFPaymentDetails[key] * 100) / 100)
                else:    
                    amount1 = HHHconf.moneyFormat.format(dict_EFPaymentDetails[key])
                amount2 = '({:0.6f})'.format(dict_EFPaymentDetails[key])
                self.findChild(QtWidgets.QLineEdit, 'Textbox_EFPaymentView_' + key).setText(amount1 + ' '*(36-len(amount1)-len(amount2)) + amount2)
            elif key == HHHconf.repayment_date:
                self.findChild(QtWidgets.QLineEdit, 'Textbox_EFPaymentView_' + key).setText(dict_EFPaymentDetails[key].strftime(HHHconf.dateFormat2 + ' %A'))
            else:
                self.findChild(QtWidgets.QLineEdit, 'Textbox_EFPaymentView_' + key).setText(str(dict_EFPaymentDetails[key]))
        self.textEdit_paymentNotes.setText(dict_EFPaymentDetails[HHHconf.payment_notes])

        # Get and display Subordinates (each subordinate displayed is LATEST updated version by TRANSID)
        df_currentSubordinates = HHHfunc.equipmentFinanceEngine.getCurrentSubordinates(selectedAgreement=self.selectedAgreement, selectedPayment=self.selectedPayment, collapseGST=True, includeHistorical=False)
        # Configure dataframe to fit table
        df_currentSubordinates.drop(['TransID', HHHconf.payment_number, HHHconf.subordinate_date, HHHconf.subordinate_created_by, HHHconf.subordinate_gst], axis=1, inplace=True)
        df_currentSubordinates.replace([None], '', inplace=True) # handles SQL converting NULL to [None]

        # calculate and display outstanding subordinates
        outstandingSubordinatesAmount = df_currentSubordinates.loc[df_currentSubordinates[HHHconf.subordinate_status].isin(HHHconf.outstandingStatuses), HHHconf.subordinate_amount].sum()
        self.findChild(QtWidgets.QLineEdit, 'Textbox_EFPaymentView_' + 'outstandingSubordinates').setText(HHHconf.moneyFormat.format(outstandingSubordinatesAmount))

        # Format columns
        df_currentSubordinates[HHHconf.subordinate_value_date] = df_currentSubordinates[HHHconf.subordinate_value_date].dt.strftime(HHHconf.dateFormat2)
        df_currentSubordinates[HHHconf.subordinate_amount] = df_currentSubordinates[HHHconf.subordinate_amount].map(HHHconf.moneyFormat.format)
        df_currentSubordinates.rename({HHHconf.subordinate_number:'Subordinate Number', HHHconf.subordinate_type:'Subordinate Type', HHHconf.subordinate_amount:'Amount', HHHconf.subordinate_status:'Status', HHHconf.subordinate_value_date:'Value Date', HHHconf.subordinate_notes:'Notes'}, axis='columns', inplace=True)
        self.EFPaymentSubordinateTable.setModel(df_currentSubordinates)

    def update_paymentDetails(self):
        # Update Payment Notes
        with HHHfunc.mainEngine.establishSqlConnection() as cnxn:
            with cnxn.cursor() as cursor:
                current_paymentNotes = str(self.textEdit_paymentNotes.toPlainText())
                if len(current_paymentNotes) == 0:
                    current_paymentNotes = None
                cursor.execute('UPDATE ' + HHHconf.db_payoutSchedule_name + ' SET ' + HHHconf.dict_toSQL_HHH_EF_Schedule[HHHconf.payment_notes] + ' = ?  WHERE ' + HHHconf.dict_toSQL_HHH_EF_Schedule[HHHconf.payment_number] + ' = ?', current_paymentNotes, self.selectedPayment)
                cnxn.commit()
        self.signal_paymentUpdated.emit({HHHconf.events_EventTime: datetime.datetime.now(), HHHconf.events_UserName: HHHconf.username_PC, HHHconf.events_Event: HHHconf.dict_eventCategories['equipmentFinance'], HHHconf.events_EventDescription: 'Updated Payment Notes | ' + self.selectedPayment + ': ' + current_paymentNotes + ' | '})
        return HHHconf.widgetEffects.flashMessage(self.heading_status, duration=1000, message='Updated Payment Notes', messageType='success')

    def EFPaymentSubordinateTable_doubleClicked(self, clickedLine):
        selectedSubordinate = str(self.EFPaymentSubordinateTable.model().index(self.EFPaymentSubordinateTable.currentIndex().row(), 0).data(QtCore.Qt.DisplayRole))
        self.openSubordinate(selectedSubordinate)

    def openSubordinate(self, subordinate=None):
        if subordinate is None:
            return
        self.childWindow = subordinateEditScreen(selectedSubordinate=subordinate, username=self.username, geometry=self.geometry(), padding={'left':0, 'top':40, 'right':0, 'bottom':260}, title='Amend Subordinate Details', updateButton='Update Subordinate', parent=self)
        self.childWindow.show()
        self.childWindow.activateWindow()
        self.childWindow.setFocus()
        self.childWindow.signal_subordinateDetailsUpdated.connect(self.get_EFPaymentDetails)
        self.childWindow.signal_subordinateDetailsUpdated.connect(lambda x: self.signal_paymentUpdated.emit(x))

    def addSubordinate(self):
        selectedPayment = str(self.selectedPayment)
        self.childWindow = subordinateAddScreen(selectedPayment=selectedPayment, username=self.username, geometry=self.geometry(), padding={'left':70, 'top':40, 'right':70, 'bottom':360}, title='Add New Subordinate', updateButton='Create Subordinate', parent=self)
        self.childWindow.show()
        self.childWindow.activateWindow()
        self.childWindow.setFocus()
        self.childWindow.signal_subordinateAdded.connect(self.get_EFPaymentDetails)
        self.childWindow.signal_subordinateAdded.connect(lambda x: self.signal_paymentUpdated.emit(x))
        self.childWindow.signal_subordinateAdded.connect(lambda x: HHHconf.widgetEffects.flashMessage(self.heading_status, duration=3000, message='Added new subordinate', messageType='success'))

class subordinateEditScreen(HHHconf.HHHWindowWidget):

    signal_subordinateDetailsUpdated = QtCore.pyqtSignal(object)

    def __init__(self, selectedSubordinate, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.selectedSubordinate = selectedSubordinate
        self.addWidgets()
        self.getSubordinateDetails(self.selectedSubordinate)

    def addWidgets(self):
        hBox_subordinateDetails = QtWidgets.QHBoxLayout()
        hBox_subordinateDetails.setContentsMargins(0, 0, 0, 0)
        hBox_subordinateDetails.setSpacing(10)
        self.vBox_contents.addLayout(hBox_subordinateDetails)

        # Current Subordinate Details
        self.currentSubordinateBox = QtWidgets.QGroupBox()
        self.currentSubordinateBox.setContentsMargins(0, 0, 0, 0)
        self.currentSubordinateBox.setStyleSheet(HHHconf.design_GroupBoxTwo)
        self.currentSubordinateBox.setMinimumWidth(160)
        hBox_subordinateDetails.addWidget(self.currentSubordinateBox)

        vBox_currentSubordinate = QtWidgets.QVBoxLayout()
        vBox_currentSubordinate.setContentsMargins(5, 5, 5, 5)
        vBox_currentSubordinate.setSpacing(2)
        self.currentSubordinateBox.setLayout(vBox_currentSubordinate)

        heading_currentSubordinate = QtWidgets.QLabel('Selected Subordinate')
        heading_currentSubordinate.setStyleSheet(HHHconf.design_textMedium)
        heading_currentSubordinate.setMaximumHeight(20)
        vBox_currentSubordinate.addWidget(heading_currentSubordinate)

        dict_currentSubordinateDetails = {
            'Subordinate Number': HHHconf.subordinate_number, 
            'Subordinate Type':HHHconf.subordinate_type, 
            'Amount': HHHconf.subordinate_amount, 
            'GST': 'subordinate_gst', 
            'Date Created':HHHconf.subordinate_date, 
            'Status':HHHconf.subordinate_status, 
            'Value Date':HHHconf.subordinate_value_date
        }
        for heading, value in dict_currentSubordinateDetails.items():
            vBox_currentDetail = QtWidgets.QVBoxLayout()
            vBox_currentDetail.setContentsMargins(0, 0, 0, 0)
            vBox_currentDetail.setSpacing(0)
            vBox_currentDetail.setAlignment(QtCore.Qt.AlignLeft)
            vBox_currentSubordinate.addLayout(vBox_currentDetail)

            heading_currentDetail = QtWidgets.QLabel(heading + ':')
            heading_currentDetail.setStyleSheet(HHHconf.design_textSmall)
            heading_currentDetail.setSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)
            vBox_currentDetail.addWidget(heading_currentDetail)

            value_currentDetail = QtWidgets.QLineEdit(objectName=value)
            value_currentDetail.setStyleSheet(HHHconf.design_editBoxTwo.replace('padding-left:20px', 'padding-left:5px'))
            value_currentDetail.setPlaceholderText('None')
            value_currentDetail.setReadOnly(True)
            vBox_currentDetail.addWidget(value_currentDetail)

        vBox_subordinateOptionsButtons = QtWidgets.QVBoxLayout()
        vBox_subordinateOptionsButtons.setContentsMargins(0, 0, 0, 0)
        vBox_subordinateOptionsButtons.setSpacing(10)
        vBox_subordinateOptionsButtons.setAlignment(QtCore.Qt.AlignCenter)
        hBox_subordinateDetails.addLayout(vBox_subordinateOptionsButtons)

        # Subordinate Actions Buttons and associated stacks
        self.dict_subordinateOptionsButtons = {
            'button_reschedule':['Reschedule Payment', HHHconf.dict_paymentStatus[HHHconf.paymentStatus_rescheduled]], 
            'button_credit':['Credit Settle', HHHconf.dict_paymentStatus[HHHconf.paymentStatus_credited]], 
            'button_received':['Payment Received', HHHconf.dict_paymentStatus[HHHconf.paymentStatus_received]]
        }
        self.subordinateOptionsButtonGroup = QtWidgets.QButtonGroup()
        self.subordinateOptionsButtonGroup.setExclusive(True)
        self.subordinateOptionsButtonGroup.buttonClicked.connect(self.displayStack)
        for button, values in self.dict_subordinateOptionsButtons.items():
            newButton = QtWidgets.QPushButton(values[0], objectName=button, autoDefault=False) 
            newButton.setFixedHeight(20)
            newButton.setCheckable(True)
            newButton.setStyleSheet(HHHconf.design_smallButtonTransparent)
            newButton.setFocusPolicy(QtCore.Qt.NoFocus)
            vBox_subordinateOptionsButtons.addWidget(newButton)
            self.subordinateOptionsButtonGroup.addButton(newButton)
 
        # Notes and Created By Section
        vBox_notes = QtWidgets.QVBoxLayout()
        vBox_notes.setContentsMargins(0, 0, 0, 0)
        vBox_notes.setSpacing(10)
        vBox_notes.setAlignment(QtCore.Qt.AlignLeft)
        vBox_subordinateOptionsButtons.addLayout(vBox_notes)

        heading_subordinateNotes = QtWidgets.QLabel('Notes')
        heading_subordinateNotes.setStyleSheet(HHHconf.design_textSmall)
        vBox_notes.addWidget(heading_subordinateNotes)
        subordinateNotes = HHHconf.HHHTextEditWidget(self, objName=HHHconf.subordinate_notes, placeholderText=HHHconf.widgetPlaceHolderText_notes, readOnly=True)
        vBox_notes.addWidget(subordinateNotes)

        heading_subordinateCreatedBy = QtWidgets.QLabel('Created By')
        heading_subordinateCreatedBy.setStyleSheet(HHHconf.design_textSmall)
        vBox_notes.addWidget(heading_subordinateCreatedBy)
        subordinateCreatedBy = QtWidgets.QLineEdit(objectName=HHHconf.subordinate_created_by)
        subordinateCreatedBy.setReadOnly(True)
        subordinateCreatedBy.setFixedHeight(20)
        subordinateCreatedBy.setStyleSheet(HHHconf.design_editBoxTwo)
        subordinateCreatedBy.setPlaceholderText('None')
        vBox_notes.addWidget(subordinateCreatedBy)

        # New Subordinate Details Stack
        self.newSubordinateStack = QtWidgets.QStackedWidget()
        hBox_subordinateDetails.addWidget(self.newSubordinateStack)

        self.newSubordinateBox_none = QtWidgets.QGroupBox(objectName='newSubordinateBox_none')
        self.newSubordinateBox_none.setStyleSheet(HHHconf.design_GroupBoxTwo)
        self.newSubordinateBox_none.setMinimumWidth(160)
        self.newSubordinateStack.addWidget(self.newSubordinateBox_none)
        self.newSubordinateStack.setCurrentWidget(self.newSubordinateBox_none)

        vBox_newSubordinate_none = QtWidgets.QVBoxLayout()
        vBox_newSubordinate_none.setContentsMargins(5, 5, 5, 5)
        vBox_newSubordinate_none.setSpacing(0)
        vBox_newSubordinate_none.setAlignment(QtCore.Qt.AlignCenter)
        self.newSubordinateBox_none.setLayout(vBox_newSubordinate_none)

        heading_newSubordinate_none = QtWidgets.QLabel('Select an action to update subordinate details - then \'Update Subordinate\' to confirm')
        heading_newSubordinate_none.setStyleSheet(HHHconf.design_textMedium.replace('AlignLeft', 'AlignCenter').replace(HHHconf.mainFontColor, 'gray'))
        heading_newSubordinate_none.setWordWrap(True)
        heading_newSubordinate_none.setSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.MinimumExpanding)  
        vBox_newSubordinate_none.addWidget(heading_newSubordinate_none)

        self.newSubordinateBox = QtWidgets.QGroupBox(objectName='newSubordinateBox')
        self.newSubordinateBox.setStyleSheet(HHHconf.design_GroupBoxTwo)
        self.newSubordinateBox.setMinimumWidth(160)
        self.newSubordinateStack.addWidget(self.newSubordinateBox)

        vBox_newSubordinate = QtWidgets.QVBoxLayout()
        vBox_newSubordinate.setContentsMargins(5, 5, 5, 5)
        vBox_newSubordinate.setSpacing(10)
        vBox_newSubordinate_none.setAlignment(QtCore.Qt.AlignTop)
        self.newSubordinateBox.setLayout(vBox_newSubordinate)

        heading_newSubordinate = QtWidgets.QLabel('Update Subordinate:')
        heading_newSubordinate.setStyleSheet(HHHconf.design_textMedium)
        heading_newSubordinate.setMaximumHeight(20)
        vBox_newSubordinate.addWidget(heading_newSubordinate)

        vBox_subordinateUpdateStatus = QtWidgets.QVBoxLayout()
        vBox_subordinateUpdateStatus.setContentsMargins(0, 0, 0, 0)
        vBox_subordinateUpdateStatus.setSpacing(0)
        vBox_subordinateUpdateStatus.setAlignment(QtCore.Qt.AlignLeft)
        vBox_newSubordinate.addLayout(vBox_subordinateUpdateStatus)

        self.heading_subordinateUpdateStatus = QtWidgets.QLabel('Status:')
        self.heading_subordinateUpdateStatus.setStyleSheet(HHHconf.design_textSmall)
        self.heading_subordinateUpdateStatus.setMaximumHeight(20)
        vBox_subordinateUpdateStatus.addWidget(self.heading_subordinateUpdateStatus)
        self.value_subordinateUpdateStatus = QtWidgets.QLineEdit('')
        self.value_subordinateUpdateStatus.setStyleSheet(HHHconf.design_editBoxTwo)
        self.value_subordinateUpdateStatus.setReadOnly(True)
        vBox_subordinateUpdateStatus.addWidget(self.value_subordinateUpdateStatus)

        vBox_subordinateUpdateValueDate = QtWidgets.QVBoxLayout()
        vBox_subordinateUpdateValueDate.setContentsMargins(0, 0, 0, 0)
        vBox_subordinateUpdateValueDate.setSpacing(0)
        vBox_subordinateUpdateValueDate.setAlignment(QtCore.Qt.AlignLeft)
        vBox_newSubordinate.addLayout(vBox_subordinateUpdateValueDate)

        self.heading_subordinateUpdateValueDate = QtWidgets.QLabel('Value Date:', objectName='heading_subordinateUpdateValueDate')
        self.heading_subordinateUpdateValueDate.setStyleSheet(HHHconf.design_textSmallError)
        self.heading_subordinateUpdateValueDate.setMaximumHeight(20)
        vBox_subordinateUpdateValueDate.addWidget(self.heading_subordinateUpdateValueDate)
        self.value_subordinateUpdateValueDate = HHHconf.calendarLineEditWidget(parent=self, objName='value_subordinateUpdateValueDate', adjustmentWidget=self.newSubordinateStack, positionTweakY=self.padding['top'])
        vBox_subordinateUpdateValueDate.addWidget(self.value_subordinateUpdateValueDate)

        vBox_updateNotes = QtWidgets.QVBoxLayout()
        vBox_updateNotes.setContentsMargins(0, 0, 0, 0)
        vBox_updateNotes.setSpacing(0)
        vBox_updateNotes.setAlignment(QtCore.Qt.AlignLeft)
        vBox_newSubordinate.addLayout(vBox_updateNotes)

        self.heading_subordinateUpdateNotes = QtWidgets.QLabel('Notes', objectName='heading_subordinateUpdateNotes')
        self.heading_subordinateUpdateNotes.setStyleSheet(HHHconf.design_textSmallError)
        vBox_updateNotes.addWidget(self.heading_subordinateUpdateNotes)

        self.value_subordinateUpdateNotes = HHHconf.HHHTextEditWidget(self, objName='value_subordinateUpdateNotes', placeholderText=HHHconf.widgetPlaceHolderText_notes)
        self.value_subordinateUpdateNotes.textChanged.connect(self.validate_updateNotes)
        vBox_updateNotes.addWidget(self.value_subordinateUpdateNotes)

        vBox_newSubordinate.addStretch(1)

        # Subordinate History List
        vBox_subordinateHistory = QtWidgets.QVBoxLayout()
        vBox_subordinateHistory.setContentsMargins(0, 0, 0, 10)
        vBox_subordinateHistory.setSpacing(0)
        self.vBox_contents.addLayout(vBox_subordinateHistory)

        heading_subordinateHistory = QtWidgets.QLabel('Subordinate History')
        heading_subordinateHistory.setStyleSheet(HHHconf.design_textMedium)
        vBox_subordinateHistory.addWidget(heading_subordinateHistory)

        self.subordinateHistoryList = QtWidgets.QListWidget()
        self.subordinateHistoryList.setFocusPolicy(QtCore.Qt.NoFocus)
        self.subordinateHistoryList.setStyleSheet(HHHconf.design_listReadOnly + HHHconf.design_tableVerticalScrollBar)
        self.subordinateHistoryList.setMinimumHeight(200)
        self.subordinateHistoryList.setMaximumHeight(200)
        self.subordinateHistoryList.currentItemChanged.connect(self.displaySelectedSubordinate)
        vBox_subordinateHistory.addWidget(self.subordinateHistoryList)

        self.button_update.clicked.connect(self.updateSubordinate)

    def getSubordinateDetails(self, selectedSubordinate):
        df_subordinateDetails = HHHfunc.equipmentFinanceEngine.getCurrentSubordinates(includeHistorical=True)
        df_subordinateDetails = df_subordinateDetails.loc[df_subordinateDetails[HHHconf.subordinate_number] == selectedSubordinate]

        self.subordinateHistoryList.clear()
        self.list_subordinateHistory = df_subordinateDetails.to_dict(orient='records')

        dict_paymentStatusDescriptors = {
                HHHconf.dict_paymentStatus[HHHconf.paymentStatus_pending]:'Created', 
                HHHconf.dict_paymentStatus[HHHconf.paymentStatus_due]:'Payment Due', 
                HHHconf.dict_paymentStatus[HHHconf.paymentStatus_uncleared]:'Direct Debit Scheduled', 
                HHHconf.dict_paymentStatus[HHHconf.paymentStatus_overdue]:'Payment Overdue', 
                HHHconf.dict_paymentStatus[HHHconf.paymentStatus_rescheduled]:'Rescheduled', 
                HHHconf.dict_paymentStatus[HHHconf.paymentStatus_received]:'Payment Received', 
                HHHconf.dict_paymentStatus[HHHconf.paymentStatus_credited]:'Transaction Credited'
        }

        for record in self.list_subordinateHistory:
            displayNotes = (str(dict_paymentStatusDescriptors[record[HHHconf.subordinate_status]]) + ': ' + str(record[HHHconf.subordinate_notes]))
            if len(displayNotes) > 70:
                displayNotes = displayNotes[:70]
            lineBreak = displayNotes.find('\n')
            if lineBreak >= 0:
                displayNotes = displayNotes[:lineBreak]
            if displayNotes != (str(dict_paymentStatusDescriptors[record[HHHconf.subordinate_status]]) + ': ' + str(record[HHHconf.subordinate_notes])):
                displayNotes = displayNotes + '...'

            listString = record[HHHconf.subordinate_date].strftime(HHHconf.dateFormat2 + ' %H:%M:%S') + ' - ' + displayNotes
            listItem = QtWidgets.QListWidgetItem(listString, self.subordinateHistoryList)
            listItem.setData(QtCore.Qt.UserRole, record) # list item holds dict values for the given subordinate
        self.subordinateHistoryList.setCurrentItem(listItem) # set to latest record
        self.subordinateHistoryList.scrollToBottom()
        
    def displaySelectedSubordinate(self, selectedSubordinateInList):
        if selectedSubordinateInList is None:
            selectedSubordinateInList = self.subordinateHistoryList.item(0)
        self.dict_itemData = selectedSubordinateInList.data(QtCore.Qt.UserRole)
        if self.dict_itemData is None:
            self.dict_itemData = dict.fromkeys(HHHconf.dict_toSQL_HHH_EF_Subordinates, None)
            self.dict_itemData['TransID'] = None
        for dataType, value in self.dict_itemData.items():
            if dataType in ['TransID', HHHconf.payment_number, HHHconf.subordinate_notes]:
                continue
            if value is None:
                self.findChild(QtWidgets.QLineEdit, dataType).clear()
            elif dataType in [HHHconf.subordinate_number, HHHconf.subordinate_type, HHHconf.subordinate_status, HHHconf.subordinate_created_by]:
                self.findChild(QtWidgets.QLineEdit, dataType).setText(value)
            elif dataType in [HHHconf.subordinate_amount, HHHconf.subordinate_gst]:
                current_item = math.ceil(value * 100) / 100
                self.findChild(QtWidgets.QLineEdit, dataType).setText(HHHconf.moneyFormat.format(current_item))
            elif dataType == HHHconf.subordinate_date:
                self.findChild(QtWidgets.QLineEdit, dataType).setText(value.strftime(HHHconf.dateFormat2 + ' %H:%M:%S'))
            elif dataType == HHHconf.subordinate_value_date:
                self.findChild(QtWidgets.QLineEdit, dataType).setText(value.strftime(HHHconf.dateFormat2))
            
        if self.dict_itemData[HHHconf.subordinate_notes] is None:
            self.findChild(QtWidgets.QTextEdit, HHHconf.subordinate_notes).clear()
        else:
            self.findChild(QtWidgets.QTextEdit, HHHconf.subordinate_notes).setText(self.dict_itemData[HHHconf.subordinate_notes])
        
    def updateSubordinate(self):
        # TODO - move sql function to HHHfunc
        for validate_status in [self.validate_valueDate(), self.validate_updateNotes()]:
            if validate_status != True:
                return

        # Add updated subordinate to subordinate db
        self.newSubordinateData = self.dict_itemData
        del self.newSubordinateData['TransID']
        self.newSubordinateData[HHHconf.subordinate_status] = self.value_subordinateUpdateStatus.text()
        self.newSubordinateData[HHHconf.subordinate_value_date] = datetime.datetime.strptime(self.value_subordinateUpdateValueDate.text(), HHHconf.dateFormat2)
        self.newSubordinateData[HHHconf.subordinate_notes] = self.value_subordinateUpdateNotes.toPlainText()
        self.newSubordinateData[HHHconf.subordinate_created_by] = self.username
        self.newSubordinateData[HHHconf.subordinate_date] = datetime.datetime.now()
        if self.newSubordinateData[HHHconf.subordinate_status] == HHHconf.dict_paymentStatus[HHHconf.paymentStatus_rescheduled]:
            if datetime.datetime.strftime(self.newSubordinateData[HHHconf.subordinate_value_date], HHHconf.dateFormat2) == datetime.date.today().strftime(HHHconf.dateFormat2):
                self.newSubordinateData[HHHconf.subordinate_status] = HHHconf.dict_paymentStatus[HHHconf.paymentStatus_due]
            elif datetime.datetime.strftime(self.newSubordinateData[HHHconf.subordinate_value_date], HHHconf.dateFormat2) < datetime.date.today().strftime(HHHconf.dateFormat2):
                self.newSubordinateData[HHHconf.subordinate_status] = HHHconf.dict_paymentStatus[HHHconf.paymentStatus_overdue]

        df_newSubordinateData = pandas.DataFrame(self.newSubordinateData, index=[0])
        df_newSubordinateData = df_newSubordinateData[list(HHHconf.dict_toSQL_HHH_EF_Subordinates.keys())].rename(HHHconf.dict_toSQL_HHH_EF_Subordinates, axis='columns', inplace=False)
        HHHfunc.equipmentFinanceEngine.commitDataframeToSQLDB(df_newSubordinateData, HHHconf.db_subordinate_name)

        # Reset view
        self.getSubordinateDetails(self.selectedSubordinate)
        self.resetStack()
        self.signal_subordinateDetailsUpdated.emit({HHHconf.events_EventTime: self.newSubordinateData[HHHconf.subordinate_date], HHHconf.events_UserName: HHHconf.username_PC, HHHconf.events_Event: HHHconf.dict_eventCategories['equipmentFinance'], HHHconf.events_EventDescription: 'Updated Subordinate | ' + str(self.selectedSubordinate)})

    def displayStack(self, buttonClicked):
        # Deselect button if it was already selected
        try:
            if self.subordinateOptionButtonSelected == str(buttonClicked.objectName()):
                self.newSubordinateStack.setCurrentWidget(self.newSubordinateBox_none)
                self.subordinateOptionsButtonGroup.setExclusive(False)
                buttonClicked.setChecked(False)
                self.subordinateOptionsButtonGroup.setExclusive(True)
                self.subordinateOptionButtonSelected = None
                return
        except:
            pass
        self.value_subordinateUpdateStatus.setText(self.dict_subordinateOptionsButtons[buttonClicked.objectName()][1])
        self.newSubordinateStack.setCurrentWidget(self.newSubordinateBox)
        self.subordinateOptionButtonSelected = str(buttonClicked.objectName())
        
    def resetStack(self):
        self.subordinateOptionsButtonGroup.setExclusive(False)
        for button in self.subordinateOptionsButtonGroup.buttons():
            button.setChecked(False)
        self.subordinateOptionsButtonGroup.setExclusive(True)
        self.newSubordinateStack.setCurrentWidget(self.newSubordinateBox_none)
        self.value_subordinateUpdateValueDate.clear()
        self.value_subordinateUpdateNotes.clear()

    def validate_valueDate(self):
        global old_format_value_start_date
        try:
            old_format_value_start_date
        except NameError:
            old_format_value_start_date = ''
        # Note, HHHconf.calendarLineEditWidget adds 'edit_' prefix to its QlineEdit widgets
        format_date = self.findChild(QtWidgets.QLineEdit, 'edit_value_subordinateUpdateValueDate').text().replace('/', '')
        if len(format_date) > len(old_format_value_start_date):    
            if len(format_date) == 1:
                output_date = format_date
            elif len(format_date) == 2:
                output_date = format_date + '/'
            elif len(format_date) == 3:
                output_date = format_date[:2] + '/' + format_date[2:]
            elif len(format_date) == 4:
                output_date = format_date[:2] + '/' + format_date [2:4] + '/'
            elif len(format_date) > 4:
                output_date = format_date[:2] + '/' + format_date [2:4] + '/' + format_date[4:8]
        else:
            cursorPosition = self.findChild(QtWidgets.QLineEdit, 'edit_value_subordinateUpdateValueDate').cursorPosition()            
            output_date = self.findChild(QtWidgets.QLineEdit, 'edit_value_subordinateUpdateValueDate').text()
        old_format_value_start_date = format_date
        self.findChild(QtWidgets.QLineEdit, 'edit_value_subordinateUpdateValueDate').setText(output_date)

        try:
            self.findChild(QtWidgets.QLineEdit, 'edit_value_subordinateUpdateValueDate').setCursorPosition(cursorPosition)
        except:
            pass

        if len(self.findChild(QtWidgets.QLineEdit, 'edit_value_subordinateUpdateValueDate').text()) > 0 and HHHfunc.is_datetime(self.findChild(QtWidgets.QLineEdit, 'edit_value_subordinateUpdateValueDate').text(), HHHconf.dateFormat2):
            self.findChild(QtWidgets.QLabel, 'heading_subordinateUpdateValueDate').setStyleSheet(HHHconf.design_textSmallSuccess)
            return True
        else:
            self.findChild(QtWidgets.QLabel, 'heading_subordinateUpdateValueDate').setStyleSheet(HHHconf.design_textSmallError)
            return HHHconf.subordinate_value_date

    def validate_updateNotes(self):
        if len(self.findChild(QtWidgets.QTextEdit, 'value_subordinateUpdateNotes').toPlainText()) > 0:
            self.findChild(QtWidgets.QLabel, 'heading_subordinateUpdateNotes').setStyleSheet(HHHconf.design_textSmallSuccess)
            return True
        else:
            self.findChild(QtWidgets.QLabel, 'heading_subordinateUpdateNotes').setStyleSheet(HHHconf.design_textSmallError)
            return HHHconf.subordinate_notes

class subordinateAddScreen(HHHconf.HHHWindowWidget):

    signal_subordinateAdded = QtCore.pyqtSignal(object)

    def __init__(self, selectedPayment, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.selectedPayment = selectedPayment
        self.feeComesWithGST = False
        self.addWidgets() 

    def addWidgets(self):
        subordinateDetailsWidgetList = [ #[object_name, widget_type, heading_text, readOnly, placeholder_text / placeholder_combobox_selection, textChanged / comboboxChanged, Qvalidator / comboBox_contents]
            [HHHconf.payment_number, QtWidgets.QLineEdit(self.selectedPayment), 'Payment Number:', True, None, None, None], 
            [HHHconf.subordinate_number, QtWidgets.QLineEdit(HHHconf.status_pendingAdd), 'Subordinate Number:', True, None, None, None], 
            [HHHconf.subordinate_status, QtWidgets.QLineEdit('Pending Create'), 'Status:', True, None, None, None],            
            [HHHconf.subordinate_type, HHHconf.HHHComboBox(parent=self, placeholderText='Please select transaction', readOnly=True), 'Subordinate Type:', True, 'Please select', self.validate_subordinate_type, sorted(list(HHHconf.dict_subordinateType.values()))], 
            [HHHconf.subordinate_amount, QtWidgets.QLineEdit(), 'Amount:', False, '$', self.validate_subordinate_amount, QtGui.QRegExpValidator(QtCore.QRegExp('^\d{0,10}(\.\d{0,2})?$'))], 
            [HHHconf.subordinate_gst, QtWidgets.QLineEdit(), 'GST:', True, '$', None, None], 
            [HHHconf.subordinate_value_date, HHHconf.calendarLineEditWidget(parent=self, objName=HHHconf.subordinate_value_date, positionTweakX=self.padding['right'], positionTweakY=self.padding['top']), 'Value Date:', False, None, self.validate_subordinate_value_date, None], 
            [HHHconf.subordinate_notes, HHHconf.HHHTextEditWidget(self), 'Notes:', False, HHHconf.widgetPlaceHolderText_notes, self.validate_subordinate_notes, None]
        ]
        for detail in subordinateDetailsWidgetList:
            hBox_currentSubordinateDetail = QtWidgets.QHBoxLayout()
            hBox_currentSubordinateDetail.setContentsMargins(0, 0, 0, 0)
            hBox_currentSubordinateDetail.setSpacing(10)
            hBox_currentSubordinateDetail.setAlignment(QtCore.Qt.AlignTop)  
            self.vBox_contents.addLayout(hBox_currentSubordinateDetail)

            currentHeading = QtWidgets.QLabel(detail[2], objectName='Heading_subAdd_' + detail[0])
            if detail[3] == False:
                currentHeading.setStyleSheet(HHHconf.design_textSmallError)
            else:                
                currentHeading.setStyleSheet(HHHconf.design_textSmall)
            currentHeading.setFixedWidth(200)
            currentHeading.setMaximumHeight(200)
            hBox_currentSubordinateDetail.addWidget(currentHeading)

            currentWidget = detail[1]
            currentWidget.setObjectName('Object_subAdd_' + detail[0])
            currentWidget.setFixedHeight(20)
            if isinstance(currentWidget, HHHconf.calendarLineEditWidget):
                currentWidget.setReadOnly(detail[3])
                if detail[5] is not None:
                    currentWidget.textChanged.connect(detail[5])
            elif isinstance(currentWidget, QtWidgets.QLineEdit):
                currentWidget.setStyleSheet(HHHconf.design_editBoxTwo)
                currentWidget.setReadOnly(detail[3])                
                if detail[4] is not None:
                    currentWidget.setPlaceholderText(detail[4])
                if detail[5] is not None:
                    currentWidget.textChanged.connect(detail[5])
                if detail[6] is not None:
                    currentWidget.setValidator(detail[6])
            elif isinstance(currentWidget, HHHconf.HHHComboBox):
                if detail[5] is not None:
                        currentWidget.currentIndexChanged.connect(detail[5])
                if detail[6] is not None:
                    for subordinateType in detail[6]:
                        currentWidget.addItem(subordinateType)
                currentWidget.setCurrentIndex(-1)                       
            elif isinstance(currentWidget, QtWidgets.QTextEdit):        
                currentWidget.setFixedHeight(200)
                currentWidget.setStyleSheet(HHHconf.design_textEdit + HHHconf.design_tableVerticalScrollBar)
                currentWidget.setReadOnly(detail[3])
                if detail[4] is not None:
                    currentWidget.setPlaceholderText(detail[4])
                if detail[5] is not None:
                    currentWidget.textChanged.connect(detail[5])
                if detail[6] is not None:
                    currentWidget.setValidator(detail[6])
            hBox_currentSubordinateDetail.addWidget(currentWidget)
        # Determines if fee comes with GST
        self.findChild(QtWidgets.QComboBox, 'Object_subAdd_' + HHHconf.subordinate_type).currentIndexChanged.connect(self.validate_gst)
        self.button_update.clicked.connect(self.addSubordinateCommit)

    def addSubordinateCommit(self):
        # TODO move sql component to HHHfunc
        for validate_status in [self.validate_subordinate_type(self.findChild(QtWidgets.QComboBox, 'Object_subAdd_' + str(HHHconf.subordinate_type)).currentIndex()), self.validate_subordinate_amount(), self.validate_subordinate_value_date(), self.validate_subordinate_notes()]:
            if validate_status != True:
                return

        # Create dictionary for new subordinate data
        dict_addSubordinateData = {}
        for child in self.findChildren(QtWidgets.QLineEdit):
            try:
                dict_addSubordinateData[HHHconf.dict_toSQL_HHH_EF_Subordinates[str(child.objectName()).replace('Object_subAdd_', '')]] = child.text()
            except:
                pass
        for child in self.findChildren(QtWidgets.QComboBox):
            dict_addSubordinateData[HHHconf.dict_toSQL_HHH_EF_Subordinates[str(child.objectName()).replace('Object_subAdd_', '')]] = child.currentText()  
        for child in self.findChildren(QtWidgets.QTextEdit):
              dict_addSubordinateData[HHHconf.dict_toSQL_HHH_EF_Subordinates[str(child.objectName()).replace('Object_subAdd_', '')]] = child.toPlainText()  
        dict_addSubordinateData[HHHconf.dict_toSQL_HHH_EF_Subordinates[HHHconf.subordinate_value_date]] = datetime.datetime.strptime(dict_addSubordinateData[HHHconf.dict_toSQL_HHH_EF_Subordinates[HHHconf.subordinate_value_date]], HHHconf.dateFormat2)
        dict_addSubordinateData[HHHconf.dict_toSQL_HHH_EF_Subordinates[HHHconf.subordinate_amount]] = float(dict_addSubordinateData[HHHconf.dict_toSQL_HHH_EF_Subordinates[HHHconf.subordinate_amount]])
        dict_addSubordinateData[HHHconf.dict_toSQL_HHH_EF_Subordinates[HHHconf.subordinate_gst]] = float(dict_addSubordinateData[HHHconf.dict_toSQL_HHH_EF_Subordinates[HHHconf.subordinate_gst]])
        if datetime.datetime.strftime(dict_addSubordinateData[HHHconf.dict_toSQL_HHH_EF_Subordinates[HHHconf.subordinate_value_date]], HHHconf.dateFormat2) == datetime.date.today().strftime(HHHconf.dateFormat2):
            dict_addSubordinateData[HHHconf.dict_toSQL_HHH_EF_Subordinates[HHHconf.subordinate_status]] = HHHconf.dict_paymentStatus[HHHconf.paymentStatus_due]
        elif datetime.datetime.strftime(dict_addSubordinateData[HHHconf.dict_toSQL_HHH_EF_Subordinates[HHHconf.subordinate_value_date]], HHHconf.dateFormat2) < datetime.date.today().strftime(HHHconf.dateFormat2):
            dict_addSubordinateData[HHHconf.dict_toSQL_HHH_EF_Subordinates[HHHconf.subordinate_status]] = HHHconf.dict_paymentStatus[HHHconf.paymentStatus_overdue]
        else:
            dict_addSubordinateData[HHHconf.dict_toSQL_HHH_EF_Subordinates[HHHconf.subordinate_status]] = HHHconf.dict_paymentStatus[HHHconf.paymentStatus_pending]
        dict_addSubordinateData[HHHconf.dict_toSQL_HHH_EF_Subordinates[HHHconf.subordinate_created_by]] = self.username
        dict_addSubordinateData[HHHconf.dict_toSQL_HHH_EF_Subordinates[HHHconf.subordinate_date]] = datetime.datetime.now()
        
        # Add subordinate data to subordinate db
        cnxn = HHHfunc.mainEngine.establishSqlConnection()
        cursor = cnxn.cursor()

        columnNames = ', '.join(list(dict_addSubordinateData.keys()))
        placeholders = ', '.join(['?'] * len(list(dict_addSubordinateData.keys())))
        Job_addSubordinate = ('INSERT INTO ' + HHHconf.db_subordinate_name + ' (' + columnNames + ') OUTPUT INSERTED.TransID VALUES (' + placeholders + ')')
        cursor.execute(Job_addSubordinate, list(dict_addSubordinateData.values()))
        newSubordinateID = cursor.fetchone()
        cnxn.commit()

        # Generate New Subordinate Number
        if newSubordinateID[0] is not None and dict_addSubordinateData[HHHconf.dict_toSQL_HHH_EF_Subordinates[HHHconf.subordinate_number]] == HHHconf.status_pendingAdd:
            cursor.execute('SELECT MAX(' + HHHconf.dict_toSQL_HHH_EF_Subordinates[HHHconf.subordinate_number] + ') FROM ' + HHHconf.db_subordinate_name + ' WHERE ' + HHHconf.dict_toSQL_HHH_EF_Subordinates[HHHconf.payment_number] + ' = ? AND ' + HHHconf.dict_toSQL_HHH_EF_Subordinates[HHHconf.subordinate_number] + ' != ?', self.selectedPayment, HHHconf.status_pendingAdd)
            maxSubordinateNumber = cursor.fetchone()
            cnxn.commit()                    
            newSubordinateIndex = int(maxSubordinateNumber[0].split('_')[2]) + 1
            newSubordinateNumber = self.selectedPayment + '_' '{0:0>3}'.format(str(newSubordinateIndex))
            cursor.execute('UPDATE ' + HHHconf.db_subordinate_name + ' SET ' + HHHconf.dict_toSQL_HHH_EF_Subordinates[HHHconf.subordinate_number] + ' = ? WHERE TransID = ' + str(newSubordinateID[0]), newSubordinateNumber)
            cnxn.commit()

        # Reset view
        self.signal_subordinateAdded.emit({HHHconf.events_EventTime: dict_addSubordinateData[HHHconf.dict_toSQL_HHH_EF_Subordinates[HHHconf.subordinate_date]], HHHconf.events_UserName: HHHconf.username_PC, HHHconf.events_Event: HHHconf.dict_eventCategories['equipmentFinance'], HHHconf.events_EventDescription: 'Added Subordinate | ' + str(newSubordinateNumber)})
        
        self.close()

    def validate_subordinate_type(self, currentIndex):
        if currentIndex != -1:
            self.findChild(QtWidgets.QLabel, 'Heading_subAdd_' + HHHconf.subordinate_type).setStyleSheet(HHHconf.design_textSmallSuccess)
            return True
        else:
            self.findChild(QtWidgets.QLabel, 'Heading_subAdd_' + HHHconf.subordinate_type).setStyleSheet(HHHconf.design_textSmallError)
            return HHHconf.subordinate_type

    def validate_subordinate_amount(self):
        currentAmount = self.findChild(QtWidgets.QLineEdit, 'Object_subAdd_' + HHHconf.subordinate_amount).text()
        # change GST displayed
        if len(currentAmount) > 0 and self.feeComesWithGST is True:
            self.findChild(QtWidgets.QLineEdit, 'Object_subAdd_' + HHHconf.subordinate_gst).setText(str('{:.2f}'.format(round(float(currentAmount) * HHHconf.gstComponent, 2))))
            self.findChild(QtWidgets.QLabel, 'Heading_subAdd_' + HHHconf.subordinate_gst).setStyleSheet(HHHconf.design_textSmallSuccess)
        else:
            self.findChild(QtWidgets.QLabel, 'Heading_subAdd_' + HHHconf.subordinate_gst).setStyleSheet(HHHconf.design_textSmall)
            self.findChild(QtWidgets.QLineEdit, 'Object_subAdd_' + HHHconf.subordinate_gst).setText('0.00')

        if len(currentAmount) > 0 and float(currentAmount) > 0:
            self.findChild(QtWidgets.QLabel, 'Heading_subAdd_' + HHHconf.subordinate_amount).setStyleSheet(HHHconf.design_textSmallSuccess)
            return True
        else:
            self.findChild(QtWidgets.QLabel, 'Heading_subAdd_' + HHHconf.subordinate_amount).setStyleSheet(HHHconf.design_textSmallError)
            return HHHconf.subordinate_amount 

    def validate_subordinate_value_date(self):
        global old_format_value_date
        try:
            old_format_value_date
        except NameError:
            old_format_value_date = ''
        format_date = self.findChild(QtWidgets.QLineEdit, 'Object_subAdd_' + HHHconf.subordinate_value_date).text().replace('/', '')
        if len(format_date) > len(old_format_value_date):
            if len(format_date) == 1:
                output_date = format_date
            elif len(format_date) == 2:
                output_date = format_date + '/'
            elif len(format_date) == 3:
                output_date = format_date[:2] + '/' + format_date[2:]
            elif len(format_date) == 4:
                output_date = format_date[:2] + '/' + format_date [2:4] + '/'
            elif len(format_date) > 4:
                output_date = format_date[:2] + '/' + format_date [2:4] + '/' + format_date[4:8]
        else:
            output_date = self.findChild(QtWidgets.QLineEdit, 'Object_subAdd_' + HHHconf.subordinate_value_date).text()
        old_format_value_date = format_date
        self.findChild(QtWidgets.QLineEdit, 'Object_subAdd_' + HHHconf.subordinate_value_date).setText(output_date)

        if len(self.findChild(QtWidgets.QLineEdit, 'Object_subAdd_' + HHHconf.subordinate_value_date).text()) > 0 and HHHfunc.is_datetime(self.findChild(QtWidgets.QLineEdit, 'Object_subAdd_' + HHHconf.subordinate_value_date).text(), HHHconf.dateFormat2) and datetime.datetime(year=1753, month=1, day=1) <= datetime.datetime.strptime(self.findChild(QtWidgets.QLineEdit, 'Object_subAdd_' + HHHconf.subordinate_value_date).text(), HHHconf.dateFormat2) <= datetime.datetime(year=9999, month=12, day=31):
            self.findChild(QtWidgets.QLabel, 'Heading_subAdd_' + HHHconf.subordinate_value_date).setStyleSheet(HHHconf.design_textSmallSuccess)
            return True
        else:
            self.findChild(QtWidgets.QLabel, 'Heading_subAdd_' + HHHconf.subordinate_value_date).setStyleSheet(HHHconf.design_textSmallError)
            return HHHconf.subordinate_value_date 

    def validate_subordinate_notes(self):
        if len(self.findChild(QtWidgets.QTextEdit, 'Object_subAdd_' + HHHconf.subordinate_notes).toPlainText()) > 0:
            self.findChild(QtWidgets.QLabel, 'Heading_subAdd_' + HHHconf.subordinate_notes).setStyleSheet(HHHconf.design_textSmallSuccess)
            return True
        else:
            self.findChild(QtWidgets.QLabel, 'Heading_subAdd_' + HHHconf.subordinate_notes).setStyleSheet(HHHconf.design_textSmallError)
            return HHHconf.subordinate_notes

    def validate_gst(self):
        if self.findChild(QtWidgets.QComboBox, 'Object_subAdd_' + HHHconf.subordinate_type).currentText() in HHHconf.list_subordinateGST:
            self.feeComesWithGST = True
        else:
            self.feeComesWithGST = False
        self.validate_subordinate_amount()