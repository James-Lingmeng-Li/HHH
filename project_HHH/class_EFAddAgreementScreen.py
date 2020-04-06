"""EF - Add New Agreement

This screen is for adding a new Equipment Finance agreement to the database.

It is mainly a big screen of empty widgets, with a lot of user validation
"""

import sys, math, pandas, datetime, re, copy

from PyQt5 import QtCore, QtGui, QtWidgets

import HHHconf, HHHfunc

class EFAddAgreementScreen(HHHconf.HHHWindowWidget):

    signal_update_EFAgreementsTable = QtCore.pyqtSignal()
    signal_agreementAdded = QtCore.pyqtSignal(object)

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)       
        self.dict_clientDetailsAdd = dict.fromkeys(HHHconf.dict_toSQL_HHH_EF_Client_Info) # used to pass data into db
        self.dict_agreementDetailsAdd = dict.fromkeys(HHHconf.dict_toSQL_HHH_EF_Agreement_Info) # used to pass data into db
        self.headingPrefix = 'heading_'
        self.editPrefix = 'edit_'
        self.previewTimer = QtCore.QTimer(timeout=self.preview_agreement)
        self.previewTimer.setSingleShot(True)    
        self.addWidgets()

    def addWidgets(self):
        hBox_heading = QtWidgets.QHBoxLayout()
        hBox_heading.setContentsMargins(5, 0, 5, 0)
        hBox_heading.setSpacing(20)
        hBox_heading.setAlignment(QtCore.Qt.AlignBottom)
        self.vBox_contents.addLayout(hBox_heading) 
        self.heading_title = QtWidgets.QLabel('Add Agreement')
        self.heading_title.setStyleSheet(HHHconf.design_textMedium)
        hBox_heading.addWidget(self.heading_title)    

        hBox_heading.addStretch(1)

        hBox_findExistingClient = QtWidgets.QHBoxLayout()
        hBox_findExistingClient.setContentsMargins(0, 0, 0, 0)
        hBox_findExistingClient.setSpacing(5)
        hBox_findExistingClient.setAlignment(QtCore.Qt.AlignBottom)
        hBox_heading.addLayout(hBox_findExistingClient) 

        vBox_spacer = QtWidgets.QVBoxLayout()
        vBox_spacer.setContentsMargins(0, 0, 0, 0)
        vBox_spacer.setSpacing(0)
        hBox_findExistingClient.addLayout(vBox_spacer)
        vBox_spacer.addSpacing(2)
        self.heading_existingClientNumber = QtWidgets.QLabel('')
        self.heading_existingClientNumber.setStyleSheet(HHHconf.design_textSmallMonoSpace.replace(HHHconf.mainFontColor, 'gray'))
        self.heading_existingClientNumber.setAlignment(QtCore.Qt.AlignBottom)
        vBox_spacer.addWidget(self.heading_existingClientNumber)  

        self.edit_findExistingClient = QtWidgets.QLineEdit('', objectName=self.editPrefix + 'findExistingClient')
        self.edit_findExistingClient.setFixedSize(250, 20)
        self.edit_findExistingClient.setStyleSheet(HHHconf.design_editBoxTwo.replace('padding-left:20px', 'padding-left:5px'))
        self.edit_findExistingClient.setProperty('required', False)
        self.edit_findExistingClient.setPlaceholderText('Link to existing client [optional]')
        self.edit_findExistingClient.textEdited.connect(self.validateExistingClient)
        self.edit_findExistingClient.editingFinished.connect(self.populateExistingClient)
        hBox_findExistingClient.addWidget(self.edit_findExistingClient)
        self.check_findExistingClient = QtWidgets.QCheckBox(objectName='check_findExistingClient')
        self.check_findExistingClient.setStyleSheet(HHHconf.design_checkboxTick)
        self.check_findExistingClient.setEnabled(False) # Read-only - if user types in existing client, this will change checkState on or off
        hBox_findExistingClient.addWidget(self.check_findExistingClient)

        # Get all clients
        df_allClients = HHHfunc.equipmentFinanceEngine.getClientDetails()
        df_allClients = df_allClients[[HHHconf.client_name, HHHconf.client_number]]
        # Find Existing Client - EditBox Completer
        self.existingClients_model = QtGui.QStandardItemModel()
        for index, row in df_allClients.iterrows():
            currentItem = QtGui.QStandardItem()
            currentItem.setText(row[HHHconf.client_name])
            currentItem.setData(row[HHHconf.client_number], QtCore.Qt.UserRole)
            self.existingClients_model.appendRow(currentItem)
        self.existingClients_completer = QtWidgets.QCompleter()
        self.existingClients_completer.setObjectName('completerPopup')
        self.existingClients_completer.setModel(self.existingClients_model)
        self.existingClients_completer.setCompletionMode(QtWidgets.QCompleter.PopupCompletion)
        self.existingClients_completer.setFilterMode(QtCore.Qt.MatchContains)   
        self.existingClients_completer.setCaseSensitivity(QtCore.Qt.CaseInsensitive)   
        self.existingClients_completer.setMaxVisibleItems(20)
        self.existingClients_completer.highlighted[QtCore.QModelIndex].connect(self.displayExistingClientNumber)
        self.existingClients_completer.activated[QtCore.QModelIndex].connect(self.populateExistingClient)
        self.edit_findExistingClient.setCompleter(self.existingClients_completer)

        # Client Details
        self.inputClientDetailsBox= QtWidgets.QGroupBox()
        self.inputClientDetailsBox.setStyleSheet(HHHconf.design_GroupBox)
        self.vBox_contents.addWidget(self.inputClientDetailsBox)

        hBox_inputClientDetails = QtWidgets.QHBoxLayout()
        hBox_inputClientDetails.setContentsMargins(5, 5, 5, 10)
        hBox_inputClientDetails.setSpacing(20)
        hBox_inputClientDetails.setAlignment(QtCore.Qt.AlignTop)   
        self.inputClientDetailsBox.setLayout(hBox_inputClientDetails) 

        vBox_inputClientDetails1 = QtWidgets.QVBoxLayout()
        vBox_inputClientDetails1.setContentsMargins(0, 0, 0, 0)
        vBox_inputClientDetails1.setSpacing(2)
        vBox_inputClientDetails1.setAlignment(QtCore.Qt.AlignTop)  
        hBox_inputClientDetails.addLayout(vBox_inputClientDetails1) 

        clientDetailsWidgetList1 = [ #[object_name, heading text, required, textEdited, editingFinished, Validator]
            [HHHconf.client_name, 'Client Name', True, self.validate_client_name, self.resetCursorPos, None],  
            [HHHconf.street_address, 'Street Address', True, None, self.resetCursorPos, None], 
            [HHHconf.suburb, 'Suburb', True, None, self.resetCursorPos, None], 
            [HHHconf.state, 'State', True, self.validate_state, self.validate_state, None], 
            [HHHconf.postcode, 'Postcode', True, self.validate_postcode, None, QtGui.QRegExpValidator(QtCore.QRegExp('^[0-9]{4}$'))]
        ]
        for clientDetail1 in clientDetailsWidgetList1:
            currentLabel = QtWidgets.QLabel(clientDetail1[1], objectName=self.headingPrefix + clientDetail1[0])
            currentLabel.setStyleSheet(HHHconf.design_textSmall)
            vBox_inputClientDetails1.addWidget(currentLabel)
            currentEdit = QtWidgets.QLineEdit('', objectName=self.editPrefix + clientDetail1[0])
            currentEdit.resize(200, 20)
            currentEdit.setStyleSheet(HHHconf.design_editBoxTwo)
            currentEdit.setProperty('required', clientDetail1[2])
            # currentEdit.textChanged.connect(self.disable_activate_client)
            currentEdit.textEdited.connect(self.removeLinkToExistingClient)            
            if clientDetail1[2] == False:
                currentEdit.setPlaceholderText('Optional')
            if clientDetail1[3] is not None:
                currentEdit.textEdited.connect(clientDetail1[3])
            if clientDetail1[4] is not None:
                currentEdit.editingFinished.connect(clientDetail1[4])
            if clientDetail1[5] is not None:
                currentEdit.setValidator(clientDetail1[5])
            currentEdit.textChanged.connect(self.updateClientDataDict)
            vBox_inputClientDetails1.addWidget(currentEdit)

        # State EditBox Completer
        self.state_model = QtCore.QStringListModel()
        self.state_model.setStringList(HHHconf.stateList)
        self.state_completer = QtWidgets.QCompleter()
        self.state_completer.setModel(self.state_model)
        self.state_completer.setCompletionMode(QtWidgets.QCompleter.PopupCompletion)
        self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.state).setCompleter(self.state_completer)

        vBox_inputClientDetails2 = QtWidgets.QVBoxLayout()
        vBox_inputClientDetails2.setContentsMargins(0, 0, 0, 0)
        vBox_inputClientDetails2.setSpacing(2)
        vBox_inputClientDetails2.setAlignment(QtCore.Qt.AlignTop)  
        hBox_inputClientDetails.addLayout(vBox_inputClientDetails2)  

        clientDetailsWidgetList2 = [ #[object_name, heading text, required, textEdited, editingFinished, Validator]
            [HHHconf.abn, 'ABN', True, self.validate_abn, None, QtGui.QRegExpValidator(QtCore.QRegExp('^[1-9][0-9]{10}$'))],  
            [HHHconf.acn, 'ACN', False, self.validate_acn, None, QtGui.QRegExpValidator(QtCore.QRegExp('^[0-9]{9}$'))], 
            [HHHconf.contact_phone, 'Phone Number', True, None, self.resetCursorPos, QtGui.QRegExpValidator(QtCore.QRegExp('^([0-9]\d*)$'))], 
            [HHHconf.contact_email, 'Email Address', True, self.validate_email, self.resetCursorPos, None], 
            [HHHconf.contact_name, 'Contact Name', True, None, self.resetCursorPos, None]
        ]
        for clientDetail2 in clientDetailsWidgetList2:
            currentLabel = QtWidgets.QLabel(clientDetail2[1], objectName=self.headingPrefix + clientDetail2[0])
            currentLabel.setStyleSheet(HHHconf.design_textSmall)
            vBox_inputClientDetails2.addWidget(currentLabel)
            currentEdit = QtWidgets.QLineEdit('', objectName=self.editPrefix + clientDetail2[0])
            currentEdit.resize(200, 20)
            currentEdit.setStyleSheet(HHHconf.design_editBoxTwo)
            currentEdit.setProperty('required', clientDetail2[2])
            # currentEdit.textChanged.connect(self.disable_activate_client)
            currentEdit.textEdited.connect(self.removeLinkToExistingClient)
            if clientDetail2[2] == False:
                currentEdit.setPlaceholderText('Optional')
            if clientDetail2[3] is not None:
                currentEdit.textChanged.connect(clientDetail2[3])
            if clientDetail2[4] is not None:
                currentEdit.editingFinished.connect(clientDetail2[4])
            if clientDetail2[5] is not None:
                currentEdit.setValidator(clientDetail2[5])
            currentEdit.textChanged.connect(self.updateClientDataDict)
            vBox_inputClientDetails2.addWidget(currentEdit)

        # Agreement Details
        self.inputAgreementValuesBox = QtWidgets.QGroupBox()
        self.inputAgreementValuesBox.setStyleSheet(HHHconf.design_GroupBox)
        self.vBox_contents.addWidget(self.inputAgreementValuesBox)

        vBox_inputAgreementValues = QtWidgets.QVBoxLayout()
        vBox_inputAgreementValues.setContentsMargins(5, 5, 5, 10)
        vBox_inputAgreementValues.setSpacing(2)
        vBox_inputAgreementValues.setAlignment(QtCore.Qt.AlignTop)  
        self.inputAgreementValuesBox.setLayout(vBox_inputAgreementValues)

        agreementDetailsWidgetList = [ #[object_name, heading text, required, textEdited, editingFinished, Validator, newline?, placeholderText]
            [HHHconf.original_balance, 'Original Balance (max = ' + HHHconf.moneyFormat.format(HHHconf.maxOriginalBalance) + ')', True, self.validate_principal_amount, None, QtGui.QRegExpValidator(QtCore.QRegExp('^([1-9]\d*)(\.\d{0,2})?$')), True, '$'], 
            [HHHconf.balloon_amount, 'Balloon Amount (max = Original Balance)', True, self.validate_balloon_amount, None, QtGui.QRegExpValidator(QtCore.QRegExp('^([0-9]\d*)(\.\d{0,2})?$')), '    ', '$'], 
            [HHHconf.interest_rate, 'Interest Rate % (max = ' + str(HHHconf.maxInterestRate) + '% p.a.)', False, self.validate_interest_rate, None, QtGui.QRegExpValidator(QtCore.QRegExp('^([0-9]{0,2})(\.\d{0,2})?$')), True, '% p.a.'], 
            [HHHconf.periodic_repayment, 'Periodic Repayment $', False, self.validate_periodic_repayment, None, QtGui.QRegExpValidator(QtCore.QRegExp('^([1-9]\d*)(\.\d{0,2})?$')), 'OR', '$'], 
            [HHHconf.periods_per_year, 'Periods Per Year', True, self.validate_periods_per_year, self.validate_periods_per_year, QtGui.QRegExpValidator(QtCore.QRegExp('^([1-9]\d*){,2}$')), True, ', '.join(HHHconf.periodsPerYearList)], 
            [HHHconf.total_periods, 'Total Periods (' + str(HHHconf.maxYears) + ' years max)', True, self.validate_total_periods, None, QtGui.QIntValidator(1, 10000), '    ', 'periods = periods per year x years'], 
            [HHHconf.settlement_date, 'Settlement Date', True, self.validate_settlement_date, None, QtGui.QRegExpValidator(QtCore.QRegExp('^\d\d/\d\d/\d\d\d\d$')), True, HHHconf.dateFormat],            
            [HHHconf.agreement_start_date, 'Agreement Start Date (First Payment Date)', True, self.validate_agreement_start_date, None, QtGui.QRegExpValidator(QtCore.QRegExp('^\d\d/\d\d/\d\d\d\d$')), '    ', HHHconf.dateFormat], 
            [HHHconf.periodic_fee, 'Periodic Fee', True, self.validate_periodic_fee, None, QtGui.QRegExpValidator(QtCore.QRegExp('^([0-9]\d*)(\.\d{0,2})?$')), True, '$'], 
            [HHHconf.account_owner, 'Account Owner', True, self.validate_owner, self.validate_owner, None, '    ', 'Select a user']
        ]
        for agreementDetail in agreementDetailsWidgetList:
            if agreementDetail[6] == True: # new hBox required for new line
                hBox_inputAgreementValues = QtWidgets.QHBoxLayout()
                hBox_inputAgreementValues.setContentsMargins(0, 0, 0, 0)
                hBox_inputAgreementValues.setSpacing(1)
                hBox_inputAgreementValues.setAlignment(QtCore.Qt.AlignTop)  
                vBox_inputAgreementValues.addLayout(hBox_inputAgreementValues) 
            else: # continue with previous hBox
                vBox_container = QtWidgets.QVBoxLayout()
                vBox_container.setContentsMargins(1, 0, 1, 0)
                vBox_container.setSpacing(0)
                vBox_container.setAlignment(QtCore.Qt.AlignBottom) 
                hBox_inputAgreementValues.addLayout(vBox_container) 
                heading_container = QtWidgets.QLabel(str(agreementDetail[6]), objectName='connector_' + agreementDetail[0])
                heading_container.setStyleSheet(HHHconf.design_textSmall.replace('color:' + HHHconf.mainFontColor, 'color: gray'))
                vBox_container.addWidget(heading_container)

            vBox_inputAgreementDetail = QtWidgets.QVBoxLayout()
            vBox_inputAgreementDetail.setContentsMargins(0, 0, 0, 0)
            vBox_inputAgreementDetail.setSpacing(0)
            vBox_inputAgreementDetail.setAlignment(QtCore.Qt.AlignTop)  
            hBox_inputAgreementValues.addLayout(vBox_inputAgreementDetail)  

            currentLabel = QtWidgets.QLabel(agreementDetail[1], objectName=self.headingPrefix + agreementDetail[0])
            currentLabel.setStyleSheet(HHHconf.design_textSmall)
            vBox_inputAgreementDetail.addWidget(currentLabel)
            if agreementDetail[0] in [HHHconf.settlement_date, HHHconf.agreement_start_date]:
                currentEdit = HHHconf.calendarLineEditWidget(parent=self.widgetFrame, objName=self.editPrefix + agreementDetail[0], adjustmentWidget=self.inputAgreementValuesBox)
            else:
                currentEdit = QtWidgets.QLineEdit('', objectName=self.editPrefix + agreementDetail[0])
            currentEdit.resize(200, 20)
            currentEdit.setStyleSheet(HHHconf.design_editBoxTwo)
            currentEdit.setProperty('required', agreementDetail[2])
            if agreementDetail[7] is not None:
                currentEdit.setPlaceholderText(agreementDetail[7])
            if agreementDetail[3] is not None:
                currentEdit.textEdited.connect(lambda sender, f=agreementDetail[3]: f(display=True))
            if agreementDetail[4] is not None:
                currentEdit.editingFinished.connect(agreementDetail[4])
            if agreementDetail[5] is not None:
                currentEdit.setValidator(agreementDetail[5])
            currentEdit.textChanged.connect(self.updateClientDataDict)
            currentEdit.textChanged.connect(self.previewTimerStart)
            vBox_inputAgreementDetail.addWidget(currentEdit)

        # Periods Per Year Completer
        self.periods_per_year_model = QtCore.QStringListModel()
        self.periods_per_year_model.setStringList(HHHconf.periodsPerYearList)
        self.periods_per_year_completer = QtWidgets.QCompleter()
        self.periods_per_year_completer.setModel(self.periods_per_year_model)
        self.periods_per_year_completer.setCompletionMode(QtWidgets.QCompleter.PopupCompletion)
        self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.periods_per_year).setCompleter(self.periods_per_year_completer)

        # Get users - TODO: NEED TO FILTER TO HHH users only
        df_loginDetails = HHHfunc.mainEngine.getLoginDetails(username=None)
        self.list_usernames = df_loginDetails[HHHconf.userAdmin_username].values.tolist()
        self.owner_model = QtCore.QStringListModel()
        self.owner_model.setStringList(self.list_usernames)
        self.owner_completer = QtWidgets.QCompleter()
        self.owner_completer.setModel(self.owner_model)
        self.owner_completer.setCaseSensitivity(QtCore.Qt.CaseInsensitive)
        self.owner_completer.setCompletionMode(QtWidgets.QCompleter.PopupCompletion)
        self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.account_owner).setCompleter(self.owner_completer)

        # Preview Table Stack 
        self.stack_EFPreviewAgreementTable = QtWidgets.QStackedWidget()
        self.vBox_contents.addWidget(self.stack_EFPreviewAgreementTable)

        self.EFPreviewAgreementTable_none = QtWidgets.QGroupBox()
        self.EFPreviewAgreementTable_none.setStyleSheet(HHHconf.design_GroupBoxTwo)
        self.stack_EFPreviewAgreementTable.addWidget(self.EFPreviewAgreementTable_none)
        self.stack_EFPreviewAgreementTable.setCurrentWidget(self.EFPreviewAgreementTable_none)
        vBox_EFPreviewAgreementTable_none = QtWidgets.QVBoxLayout()
        vBox_EFPreviewAgreementTable_none.setContentsMargins(5, 5, 5, 5)
        vBox_EFPreviewAgreementTable_none.setSpacing(0)
        vBox_EFPreviewAgreementTable_none.setAlignment(QtCore.Qt.AlignCenter)
        self.EFPreviewAgreementTable_none.setLayout(vBox_EFPreviewAgreementTable_none)
        heading_EFPreviewAgreementTable_none = QtWidgets.QLabel('Complete all agreement values above (client details optional) for a payout preview')
        heading_EFPreviewAgreementTable_none.setStyleSheet(HHHconf.design_textLarge.replace('AlignLeft', 'AlignCenter').replace(HHHconf.mainFontColor, 'gray'))
        heading_EFPreviewAgreementTable_none.setWordWrap(True)
        heading_EFPreviewAgreementTable_none.setSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.MinimumExpanding)  
        vBox_EFPreviewAgreementTable_none.addWidget(heading_EFPreviewAgreementTable_none)

        self.EFPreviewAgreementTable = HHHconf.HHHTableWidget(cornerWidget='copy')
        self.EFPreviewAgreementTable.setMinimumHeight(200) 
        self.stack_EFPreviewAgreementTable.addWidget(self.EFPreviewAgreementTable)

        # MTX Numbers Section
        hBox_mtxNumbers = QtWidgets.QHBoxLayout()
        hBox_mtxNumbers.setContentsMargins(0, 10, 0, 0)
        hBox_mtxNumbers.setSpacing(20)
        hBox_mtxNumbers.setAlignment(QtCore.Qt.AlignTop)   
        self.vBox_contents.addLayout(hBox_mtxNumbers) 

        clientDetailsWidgetList3 = [ #[object_name, heading text, required, textEdited, editingFinished, Validator]
            [HHHconf.mtx_number, 'Mtx Buyer Account', False, self.validate_mtx, None, QtGui.QRegExpValidator(QtCore.QRegExp('^[0-9]{10}$'))], 
            [HHHconf.bsb, 'Bank BSB', False, self.validate_bankBSB, None, QtGui.QRegExpValidator(QtCore.QRegExp('^\d{3}-\d{3}$'))], 
            [HHHconf.acc, 'Bank ACC', False, self.validate_bankACC, None, QtGui.QRegExpValidator(QtCore.QRegExp('^\d{0,10}$'))]
        ]
        for detail in clientDetailsWidgetList3:
            vBox_mtxNumberBuyer = QtWidgets.QVBoxLayout()
            vBox_mtxNumberBuyer.setContentsMargins(0, 0, 0, 0)
            vBox_mtxNumberBuyer.setSpacing(0)
            vBox_mtxNumberBuyer.setAlignment(QtCore.Qt.AlignTop)   
            hBox_mtxNumbers.addLayout(vBox_mtxNumberBuyer)

            currentLabel = QtWidgets.QLabel(detail[1], objectName=self.headingPrefix + detail[0])
            currentLabel.setStyleSheet(HHHconf.design_textSmall)
            vBox_mtxNumberBuyer.addWidget(currentLabel)
            currentEdit = QtWidgets.QLineEdit('', objectName=self.editPrefix + detail[0])
            currentEdit.resize(200, 20)
            currentEdit.setStyleSheet(HHHconf.design_editBoxTwo)
            currentEdit.setProperty('required', detail[2])
            if detail[3] is not None:
                currentEdit.textEdited.connect(detail[3])
            if detail[4] is not None:
                currentEdit.editingFinished.connect(detail[4])
            if detail[5] is not None:
                currentEdit.setValidator(detail[5])
            currentEdit.editingFinished.connect(self.updateClientDataDict)
            vBox_mtxNumberBuyer.addWidget(currentEdit) 
        self.button_update.clicked.connect(self.activate_client)
           
    def validateExistingClient(self):
        # If user types in matching text, return the first match
        matchingClientsList = self.existingClients_model.findItems(self.sender().text(), flags=QtCore.Qt.MatchExactly)
        if len(matchingClientsList) == 0:
            self.heading_existingClientNumber.setText('')
            self.check_findExistingClient.setChecked(False)
        else:
            self.heading_existingClientNumber.setText(matchingClientsList[0].data(QtCore.Qt.UserRole))
            self.check_findExistingClient.setChecked(True)

    def populateExistingClient(self):
        self.edit_findExistingClient.setCursorPosition(0)
        if self.check_findExistingClient.isChecked():
            # get client details from db
            df_clientDetails = HHHfunc.equipmentFinanceEngine.getClientDetails([self.heading_existingClientNumber.text()])
            for col in df_clientDetails:
                try:
                    self.findChild(QtWidgets.QLineEdit, self.editPrefix + col).setText(df_clientDetails.at[0, col])
                except:
                    continue
            # Add client Number to dictionary that will be saved to db if activated
            self.dict_agreementDetailsAdd[HHHconf.client_number] = str(self.heading_existingClientNumber.text())
        else:
            # Clear client Number from dict
            self.dict_agreementDetailsAdd[HHHconf.client_number] = None
            self.edit_findExistingClient.clear()
    
    def displayExistingClientNumber(self, modelIndex):
        clientNumber = modelIndex.data(QtCore.Qt.UserRole)
        self.heading_existingClientNumber.setText(clientNumber)
        self.check_findExistingClient.setChecked(True)
        self.edit_findExistingClient.setCursorPosition(0)

    def removeLinkToExistingClient(self):
        self.heading_existingClientNumber.setText('')
        self.edit_findExistingClient.clear()
        self.check_findExistingClient.setChecked(False)

    def updateClientDataDict(self):
        objName = self.sender().objectName().replace(self.editPrefix, '')
        if objName in self.dict_clientDetailsAdd:
            self.dict_clientDetailsAdd[objName] = self.sender().text()
        elif objName in self.dict_agreementDetailsAdd:
            self.dict_agreementDetailsAdd[objName] = self.sender().text()

    def previewTimerStart(self):
        # Validate required fields
        for widget in self.inputAgreementValuesBox.findChildren(QtWidgets.QLineEdit):
            if widget.property('required') and len(widget.text()) == 0:
                self.df_payout_schedule = []
                self.previewTimer.stop()
                self.stack_EFPreviewAgreementTable.setCurrentWidget(self.EFPreviewAgreementTable_none)# Hide table
                if self.EFPreviewAgreementTable.model():
                    self.EFPreviewAgreementTable.model().deleteLater()
                self.disable_activate_client()
                return
        # Validate editBox rules and calculation rules
        for f in [self.validate_principal_amount, self.validate_balloon_amount, self.validate_interest_rate, self.validate_periodic_repayment, self.validate_periods_per_year, self.validate_total_periods, self.validate_settlement_date, self.validate_agreement_start_date, self.validate_periodic_fee, self.validate_owner, self.validate_interest_rate_or_periodic_repayment]:
            if f(display=False) != True:
                self.df_payout_schedule = []                
                self.previewTimer.stop()
                self.stack_EFPreviewAgreementTable.setCurrentWidget(self.EFPreviewAgreementTable_none) # Hide table                
                if self.EFPreviewAgreementTable.model():
                    self.EFPreviewAgreementTable.model().deleteLater()
                self.disable_activate_client()
                return  
        
        self.stack_EFPreviewAgreementTable.setCurrentWidget(self.EFPreviewAgreementTable)  # Show table
        self.previewTimer.start(300)
        # necessary in case user overrides dict data right when timer expires and executes preview_agreement, thereby sending potentially bad data to the engine.
        self.dict_agreementDetailsCopy = copy.deepcopy(self.dict_agreementDetailsAdd)

        HHHconf.widgetEffects.flashMessage(self.heading_status, duration=1000, message='Loading Payout Schedule Preview...', messageType='warning')

    def preview_agreement(self):
        # Get formatted data as well as calculated payout schedule
        self.dict_agreementDetailsAddFormatted, self.df_payout_schedule, self.periodicRepayment = HHHfunc.equipmentFinanceEngine.createPayoutSchedulePreview(self.dict_agreementDetailsCopy)

        # Clear error text format
        for widget in self.inputAgreementValuesBox.findChildren(QtWidgets.QLabel):
            if 'connector_' in widget.objectName():
                continue
            widget.setStyleSheet(HHHconf.design_textSmall)

        # Display payout Schedule
        self.df_payoutPreview = copy.deepcopy(self.df_payout_schedule)

        for col in self.df_payoutPreview:
            if col == 'currentPeriod':
                continue
            elif col == HHHconf.repayment_date:
                self.df_payoutPreview[col] = self.df_payoutPreview[col].dt.strftime(HHHconf.dateFormat2)
            elif col == HHHconf.periodic_repayment:
                self.df_payoutPreview[col] = self.df_payoutPreview[col].map(lambda x: HHHconf.moneyFormat.format(math.ceil(abs(x) * 100) / 100))
            else:
                self.df_payoutPreview[col] = self.df_payoutPreview[col].map(HHHconf.moneyFormat.format)
        self.df_payoutPreview.rename({'currentPeriod':'#', HHHconf.repayment_date:'Date', HHHconf.opening_balance:'Opening Balance', HHHconf.interest_component:'Interest', HHHconf.principal_component:'Principal', HHHconf.periodic_repayment:'Repayment', HHHconf.closing_balance:'Closing Balance'}, axis='columns', inplace=True)
        self.EFPreviewAgreementTable.setModel(self.df_payoutPreview)

        isValid = True # keep here for future further validation code

        # End Generate Agreement Preview
        if isValid is False:
            self.disable_activate_client()
            HHHconf.widgetEffects.flashMessage(self.heading_status, duration=3000, message='*Agreement preview generated with errors.', messageType='error')
        else:    
            self.enable_activate_client()        
            HHHconf.widgetEffects.flashMessage(self.heading_status, duration=3000, message='Agreement preview generated successfully.', messageType='success')

    def disable_activate_client(self, message='*Invalid agreement details detected'):
        self.activate_client_enabled = False
        self.activate_client_status_message = message

    def enable_activate_client(self, message='Activating client...'):
        self.activate_client_enabled = True
        self.activate_client_status_message = message 

    def activate_client(self):
        # Validate required client details
        for widget in self.inputClientDetailsBox.findChildren(QtWidgets.QLineEdit):
            if widget.property('required') and len(widget.text()) == 0:
                HHHconf.widgetEffects.flashMessage(self.heading_status, duration=3000, message='*Missing client details detected', messageType='error')
                return

        # Validate client details fields 
        for f in [self.validate_client_name, self.validate_state, self.validate_postcode, self.validate_abn, self.validate_acn, self.validate_email]:
            result = f(display=True)
            if result != True:
                HHHconf.widgetEffects.flashMessage(self.heading_status, duration=3000, message='*Invalid input detected: ' + str(result), messageType='error')
                return
        
        # Validate agreement preview
        if self.activate_client_enabled != True:
            HHHconf.widgetEffects.flashMessage(self.heading_status, duration=3000, message=self.activate_client_status_message, messageType='error')
            return
        else:
            HHHconf.widgetEffects.flashMessage(self.heading_status, duration=3000, message=self.activate_client_status_message, messageType='success')     

        # Check if MTX Number and Bank details are entered and valid
        if len(self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.mtx_number).text()) == 0 or len(self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.bsb).text()) == 0 or len(self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.acc).text()) == 0:
            self.heading_status.setStyleSheet(HHHconf.design_textSmallError)
            self.heading_status.setText('*Mtx/Bank required')
            return
        for validate_status in [self.validate_mtx(self.editPrefix + HHHconf.mtx_number), self.validate_bankBSB(), self.validate_bankACC()]:
            if validate_status != True:
                self.heading_status.setStyleSheet(HHHconf.design_textSmallError)
                self.heading_status.setText('*Invalid input detected: ' + str(validate_status))
                return

        # Get MTX Number and Bank Details into agreement details dict
        self.dict_agreementDetailsAddFormatted[HHHconf.mtx_number] = self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.mtx_number).text()
        self.dict_agreementDetailsAddFormatted[HHHconf.bsb] = self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.bsb).text()
        self.dict_agreementDetailsAddFormatted[HHHconf.acc] = self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.acc).text()

        # If new client, add client details to clientInfo table
        if self.dict_agreementDetailsAddFormatted[HHHconf.client_number] is None:
            self.dict_agreementDetailsAddFormatted[HHHconf.client_number] = HHHfunc.equipmentFinanceEngine.createNewClient(self.dict_clientDetailsAdd)
            # self.signal_update_EFAgreementsTable.emit('Added New Client - ' + self.dict_agreementDetailsAddFormatted[HHHconf.client_number] + ' ' + self.dict_clientDetailsAdd[HHHconf.client_name])

        # Add in agreement status = active
        self.dict_agreementDetailsAddFormatted[HHHconf.agreement_status] = HHHconf.dict_agreementStatus[HHHconf.agreementStatus_active]

        # Add Agreement Details to db
        self.dict_agreementDetailsAddFormatted[HHHconf.agreement_number] = HHHfunc.equipmentFinanceEngine.createAgreement(self.dict_agreementDetailsAddFormatted)
        
        # Configure df to match sql table
        self.df_payout_schedule[HHHconf.client_name] = self.dict_clientDetailsAdd[HHHconf.client_name]
        self.df_payout_schedule[HHHconf.agreement_number] = self.dict_agreementDetailsAddFormatted[HHHconf.agreement_number]
        self.df_payout_schedule[HHHconf.mtx_number] = self.dict_agreementDetailsAdd[HHHconf.mtx_number]
        self.df_payout_schedule[HHHconf.payment_status] = HHHconf.dict_paymentStatus[HHHconf.paymentStatus_pending]
        self.df_payout_schedule[HHHconf.completion_date] = None
        self.df_payout_schedule[HHHconf.payment_notes] = None
        self.df_payout_schedule[HHHconf.payment_number] = self.df_payout_schedule[HHHconf.agreement_number].map(str) + '_' + self.df_payout_schedule['currentPeriod'].map(str).str.pad(width=4, fillchar='0')

        # Reorder columns. Drop current_period as info is now stored in payment_number
        df_EFAdd_payout_schedule = self.df_payout_schedule[list(HHHconf.dict_toSQL_HHH_EF_Schedule.keys())]
        df_EFAdd_payout_schedule = df_EFAdd_payout_schedule.rename(HHHconf.dict_toSQL_HHH_EF_Schedule, axis='columns', inplace=False)

        # Add Payout Schedule to db
        HHHfunc.equipmentFinanceEngine.commitDataframeToSQLDB(df_EFAdd_payout_schedule, HHHconf.db_payoutSchedule_name)

        # Add initial subordinate transactions to db
        df_EFAdd_subordinates = pandas.DataFrame()
        df_EFAdd_subordinates[HHHconf.payment_number] = self.df_payout_schedule[HHHconf.payment_number]
        df_EFAdd_subordinates[HHHconf.subordinate_number] = self.df_payout_schedule[HHHconf.payment_number] + '_000'
        df_EFAdd_subordinates[HHHconf.subordinate_date] = datetime.datetime.now()
        df_EFAdd_subordinates[HHHconf.subordinate_value_date] = self.df_payout_schedule[HHHconf.repayment_date]
        df_EFAdd_subordinates[HHHconf.subordinate_type] = HHHconf.dict_subordinateType[HHHconf.subordinateType_repayment]
        df_EFAdd_subordinates[HHHconf.subordinate_amount] = math.ceil(self.periodicRepayment * 100) / 100 # Round up to nearest cent (subordinate transactions are the actual amounts scheduled for direct debit) Value is extracted when payout schedule was generated
        df_EFAdd_subordinates[HHHconf.subordinate_gst] = 0
        df_EFAdd_subordinates[HHHconf.subordinate_status] = HHHconf.dict_paymentStatus[HHHconf.paymentStatus_pending]
        df_EFAdd_subordinates[HHHconf.subordinate_notes] = HHHconf.subordinateNotes_systemGenerated
        df_EFAdd_subordinates[HHHconf.subordinate_created_by] = self.username

        # Add periodic fee subordinates if periodic fee set to > 0
        if self.dict_agreementDetailsAddFormatted[HHHconf.periodic_fee] > 0:
            df_EFAdd_periodicFees = copy.deepcopy(df_EFAdd_subordinates)
            df_EFAdd_periodicFees[HHHconf.subordinate_number] = self.df_payout_schedule[HHHconf.payment_number] + '_001'
            df_EFAdd_periodicFees[HHHconf.subordinate_type] = HHHconf.dict_subordinateType[HHHconf.subordinateType_periodic]
            df_EFAdd_periodicFees[HHHconf.subordinate_amount] = float(self.dict_agreementDetailsAddFormatted[HHHconf.periodic_fee])

            df_EFAdd_subordinates = pandas.concat([df_EFAdd_subordinates, df_EFAdd_periodicFees])

        HHHfunc.equipmentFinanceEngine.addGSTtoTransactions(df_EFAdd_subordinates, HHHconf.subordinate_type, HHHconf.subordinate_amount, HHHconf.subordinate_gst)

        df_EFAdd_subordinates.rename(HHHconf.dict_toSQL_HHH_EF_Subordinates, axis='columns', inplace=True)
        HHHfunc.equipmentFinanceEngine.commitDataframeToSQLDB(df_EFAdd_subordinates, HHHconf.db_subordinate_name)

        self.signal_update_EFAgreementsTable.emit()
        self.signal_agreementAdded.emit({HHHconf.events_EventTime: datetime.datetime.now(), HHHconf.events_UserName: HHHconf.username_PC, HHHconf.events_Event: HHHconf.dict_eventCategories['equipmentFinance'], HHHconf.events_EventDescription: str('Added New Agreement - ' + self.dict_agreementDetailsAddFormatted[HHHconf.agreement_number] + ' ' + self.dict_clientDetailsAdd[HHHconf.client_name])})
        self.signal_update_EFAgreementsTable.emit()
        self.close()
        # NOTIFY NEW ACCOUNT MAANAGER ON THEIR NEW ACCOUNT - TODOTODO

    def validate_client_name(self, display=True): # Currently not used
        if len(self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.client_name).text()) < 0:
            self.findChild(QtWidgets.QLabel, self.headingPrefix + HHHconf.client_name).setStyleSheet(HHHconf.design_textSmallError)
            return HHHconf.client_name
        self.findChild(QtWidgets.QLabel, self.headingPrefix + HHHconf.client_name).setStyleSheet(HHHconf.design_textSmall)
        return True

    def validate_state(self, text='', display=True):
        if text != text.upper():
            self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.state).setText(text.upper())
        if len(self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.state).text()) > 0 and self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.state).text() not in HHHconf.stateList:
            self.findChild(QtWidgets.QLabel, self.headingPrefix + HHHconf.state).setStyleSheet(HHHconf.design_textSmallError)
            return HHHconf.state
        else:
            self.findChild(QtWidgets.QLabel, self.headingPrefix + HHHconf.state).setStyleSheet(HHHconf.design_textSmall)
            return True
    
    def validate_postcode(self, display=True):
        if re.match(r'^[0-9]{4}$', self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.postcode).text()) or len(self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.postcode).text()) == 0:
            self.findChild(QtWidgets.QLabel, self.headingPrefix + HHHconf.postcode).setStyleSheet(HHHconf.design_textSmall)
            return True          
        else:
            self.findChild(QtWidgets.QLabel, self.headingPrefix + HHHconf.postcode).setStyleSheet(HHHconf.design_textSmallError)
            return HHHconf.postcode
        
    def validate_abn(self, display=True):
        abn = self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.abn).text()
        if HHHfunc.is_ABN(abn):
            self.findChild(QtWidgets.QLabel, self.headingPrefix + HHHconf.abn).setStyleSheet(HHHconf.design_textSmallSuccess)
            return True   
        elif len(self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.abn).text()) == 0:
            self.findChild(QtWidgets.QLabel, self.headingPrefix + HHHconf.abn).setStyleSheet(HHHconf.design_textSmall)
            return True   
        else:
            self.findChild(QtWidgets.QLabel, self.headingPrefix + HHHconf.abn).setStyleSheet(HHHconf.design_textSmallError)
            return HHHconf.abn
    
    def validate_acn(self, display=True):
        acn = self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.acn).text()
        if HHHfunc.is_ACN(acn):
            self.findChild(QtWidgets.QLabel, self.headingPrefix + HHHconf.acn).setStyleSheet(HHHconf.design_textSmallSuccess)
            return True 
        elif len(self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.acn).text()) == 0:
            self.findChild(QtWidgets.QLabel, self.headingPrefix + HHHconf.acn).setStyleSheet(HHHconf.design_textSmall)
            return True   
        else:
            self.findChild(QtWidgets.QLabel, self.headingPrefix + HHHconf.acn).setStyleSheet(HHHconf.design_textSmallError)
            return HHHconf.acn

    def validate_email(self, display=True):
        if re.match(r'[^@]+@[^@]+\.[^@]+', self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.contact_email).text()) or len(self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.contact_email).text()) == 0:
            self.findChild(QtWidgets.QLabel, self.headingPrefix + HHHconf.contact_email).setStyleSheet(HHHconf.design_textSmall)
            return True
        else:
            self.findChild(QtWidgets.QLabel, self.headingPrefix + HHHconf.contact_email).setStyleSheet(HHHconf.design_textSmallError)
            return HHHconf.contact_email

    def validate_principal_amount(self, display=True):
        principalAmount = self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.original_balance).text()
        if HHHfunc.is_number(principalAmount):
            if float(principalAmount) > HHHconf.maxOriginalBalance:
                if display is True:
                    self.findChild(QtWidgets.QLabel, self.headingPrefix + HHHconf.original_balance).setStyleSheet(HHHconf.design_textSmallError)
                return HHHconf.original_balance
        self.findChild(QtWidgets.QLabel, self.headingPrefix + HHHconf.original_balance).setStyleSheet(HHHconf.design_textSmall)
        self.validate_balloon_amount()
        return True        

    def validate_balloon_amount(self, display=True):
        balloonAmount = self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.balloon_amount).text()

        if len(self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.original_balance).text()) > 0 and len(balloonAmount) > 0:
            if float(balloonAmount) > float(self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.original_balance).text()):
                if display is True:
                    self.findChild(QtWidgets.QLabel, self.headingPrefix + HHHconf.balloon_amount).setStyleSheet(HHHconf.design_textSmallError)
                return HHHconf.balloon_amount
        self.findChild(QtWidgets.QLabel, self.headingPrefix + HHHconf.balloon_amount).setStyleSheet(HHHconf.design_textSmall)
        return True
        

    def validate_interest_rate(self, display=True):
        self.findChild(QtWidgets.QLabel, self.headingPrefix + HHHconf.periodic_repayment).setStyleSheet(HHHconf.design_textSmall)        
        if len(self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.periodic_repayment).text()) > 0 and len(self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.interest_rate).text()) > 0:
            self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.periodic_repayment).clear()
        interestRate = self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.interest_rate).text()
        if len(interestRate) > 0:
            self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.periodic_repayment).setReadOnly(True)
            self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.periodic_repayment).setFocusPolicy(QtCore.Qt.NoFocus)
            self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.periodic_repayment).setPlaceholderText('Interest Rate Detected')
            self.findChild(QtWidgets.QLabel, 'connector_' + HHHconf.periodic_repayment).setText('    ')
        else:
            self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.periodic_repayment).setReadOnly(False)
            self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.periodic_repayment).setFocusPolicy(QtCore.Qt.StrongFocus)
            self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.periodic_repayment).setPlaceholderText('')
            if display is True:
                self.findChild(QtWidgets.QLabel, 'connector_' + HHHconf.periodic_repayment).setText('OR')

        if HHHfunc.is_number(interestRate):
            if float(interestRate) > HHHconf.maxInterestRate:
                self.findChild(QtWidgets.QLabel, self.headingPrefix + HHHconf.interest_rate).setStyleSheet(HHHconf.design_textSmallError)
                return HHHconf.interest_rate + ' greater than max'
        self.findChild(QtWidgets.QLabel, self.headingPrefix + HHHconf.interest_rate).setStyleSheet(HHHconf.design_textSmall)
        return True        

    def validate_periodic_repayment(self, display=True):
        self.findChild(QtWidgets.QLabel, self.headingPrefix + HHHconf.interest_rate).setStyleSheet(HHHconf.design_textSmall)        
        if len(self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.periodic_repayment).text()) > 0 and len(self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.interest_rate).text()) > 0:
            self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.interest_rate).clear()
        periodicRepayment = self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.periodic_repayment).text()
        if len(periodicRepayment) > 0:
            self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.interest_rate).setReadOnly(True)
            self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.interest_rate).setFocusPolicy(QtCore.Qt.NoFocus)
            self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.interest_rate).setPlaceholderText('Periodic Repayment Detected')
            self.findChild(QtWidgets.QLabel, 'connector_' + HHHconf.periodic_repayment).setText('    ')
        else:
            self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.interest_rate).setReadOnly(False)
            self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.interest_rate).setFocusPolicy(QtCore.Qt.StrongFocus)
            self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.interest_rate).setPlaceholderText('')
            if display is True:
                self.findChild(QtWidgets.QLabel, 'connector_' + HHHconf.periodic_repayment).setText('OR')

        self.findChild(QtWidgets.QLabel, self.headingPrefix + HHHconf.periodic_repayment).setStyleSheet(HHHconf.design_textSmall)
        return True        

    def validate_interest_rate_or_periodic_repayment(self, display=True):
        if len(self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.interest_rate).text()) == 0 and len(self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.periodic_repayment).text()) == 0:
            self.findChild(QtWidgets.QLabel, self.headingPrefix + HHHconf.interest_rate).setStyleSheet(HHHconf.design_textSmallError)
            self.findChild(QtWidgets.QLabel, self.headingPrefix + HHHconf.periodic_repayment).setStyleSheet(HHHconf.design_textSmallError)
            self.findChild(QtWidgets.QLabel, 'connector_' + HHHconf.periodic_repayment).setText('OR')
            return 'Please enter either Interest Rate or Periodic Repayment'
        return True

    def validate_periods_per_year(self, display=True):
        if len(self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.periods_per_year).text()) > 0 and self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.periods_per_year).text() not in HHHconf.periodsPerYearList:
            if display is True:
                self.findChild(QtWidgets.QLabel, self.headingPrefix + HHHconf.periods_per_year).setStyleSheet(HHHconf.design_textSmallError)
            return HHHconf.periods_per_year
        else:
            self.findChild(QtWidgets.QLabel, self.headingPrefix + HHHconf.periods_per_year).setStyleSheet(HHHconf.design_textSmall)
            return True
    
    def validate_total_periods(self, display=True):
        if len(self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.periods_per_year).text()) > 0 and len(self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.total_periods).text()) > 0 and int(self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.total_periods).text()) / int(self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.periods_per_year).text()) > HHHconf.maxYears:
            if display is True:
                self.findChild(QtWidgets.QLabel, self.headingPrefix + HHHconf.total_periods).setStyleSheet(HHHconf.design_textSmallError)
            return 'Total Periods exceeds '  + str(HHHconf.maxYears) + ' years.'
        # elif len(self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.periods_per_year).text()) > 0 and len(self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.total_periods).text()) > 0:
        #     self.findChild(QtWidgets.QLabel, self.headingPrefix + HHHconf.total_periods).setStyleSheet(HHHconf.design_textSmallError)
        #     return HHHconf.total_periods
        else:
            self.findChild(QtWidgets.QLabel, self.headingPrefix + HHHconf.total_periods).setStyleSheet(HHHconf.design_textSmall)
            return True

    def validate_agreement_start_date(self, display=True):
        global old_format_agreement_start_date
        try:
            old_format_agreement_start_date
        except NameError:
            old_format_agreement_start_date = ''
        format_date = self.findChild(HHHconf.calendarLineEditWidget, 'edit_' + self.editPrefix + HHHconf.agreement_start_date).text().replace('/', '')
        if len(format_date) > len(old_format_agreement_start_date):
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
            output_date = self.findChild(HHHconf.calendarLineEditWidget, 'edit_' + self.editPrefix + HHHconf.agreement_start_date).text()
        old_format_agreement_start_date = format_date
        self.findChild(HHHconf.calendarLineEditWidget, 'edit_' + self.editPrefix + HHHconf.agreement_start_date).setText(output_date)

        if len(self.findChild(HHHconf.calendarLineEditWidget, 'edit_' + self.editPrefix + HHHconf.agreement_start_date).text()) > 0 and not HHHfunc.is_datetime(self.findChild(HHHconf.calendarLineEditWidget, 'edit_' + self.editPrefix + HHHconf.agreement_start_date).text(), HHHconf.dateFormat2):
            if display is True:
                self.findChild(QtWidgets.QLabel, self.headingPrefix + HHHconf.agreement_start_date).setStyleSheet(HHHconf.design_textSmallError)
            return HHHconf.agreement_start_date
        else:
            self.findChild(QtWidgets.QLabel, self.headingPrefix + HHHconf.agreement_start_date).setStyleSheet(HHHconf.design_textSmall)
            return True     

    def validate_settlement_date(self, display=True):
        global old_format_settlement_date
        try:
            old_format_settlement_date
        except NameError:
            old_format_settlement_date = ''
        format_date = self.findChild(HHHconf.calendarLineEditWidget, 'edit_' + self.editPrefix + HHHconf.settlement_date).text().replace('/', '')
        if len(format_date) > len(old_format_settlement_date):
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
            output_date = self.findChild(HHHconf.calendarLineEditWidget, 'edit_' + self.editPrefix + HHHconf.settlement_date).text()
        old_format_settlement_date = format_date
        self.findChild(HHHconf.calendarLineEditWidget, 'edit_' + self.editPrefix + HHHconf.settlement_date).setText(output_date)

        if len(self.findChild(HHHconf.calendarLineEditWidget, 'edit_' + self.editPrefix + HHHconf.settlement_date).text()) > 0 and not HHHfunc.is_datetime(self.findChild(HHHconf.calendarLineEditWidget, 'edit_' + self.editPrefix + HHHconf.settlement_date).text(), HHHconf.dateFormat2):
            if display is True:
                self.findChild(QtWidgets.QLabel, self.headingPrefix + HHHconf.settlement_date).setStyleSheet(HHHconf.design_textSmallError)
            return HHHconf.settlement_date
        else:
            self.findChild(QtWidgets.QLabel, self.headingPrefix + HHHconf.settlement_date).setStyleSheet(HHHconf.design_textSmall)
            return True 

    def validate_periodic_fee(self, display=True):
        periodicFeeAmount = self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.periodic_fee).text()
        if len(periodicFeeAmount) == 0:
            if display is True:
                self.findChild(QtWidgets.QLabel, self.headingPrefix + HHHconf.periodic_fee).setStyleSheet(HHHconf.design_textSmallError)
            return HHHconf.periodic_fee
        self.findChild(QtWidgets.QLabel, self.headingPrefix + HHHconf.periodic_fee).setStyleSheet(HHHconf.design_textSmall)
        return True

    def validate_owner(self, display=True):
        if len(self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.account_owner).text()) > 0 and self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.account_owner).text() not in self.list_usernames:
            self.findChild(QtWidgets.QLabel, self.headingPrefix + HHHconf.account_owner).setStyleSheet(HHHconf.design_textSmallError)
            return HHHconf.account_owner
        else:
            self.findChild(QtWidgets.QLabel, self.headingPrefix + HHHconf.account_owner).setStyleSheet(HHHconf.design_textSmall)
            return True

    def validate_mtx(self, obj=None, display=True):
        textboxName = self.sender().objectName()
        if len(textboxName) == 0:
            textboxName = str(obj)
        headingName = textboxName.replace(self.editPrefix, self.headingPrefix)
        objName = textboxName.replace(self.editPrefix, '')
        if re.match(r'^[0-9]{10}$', self.findChild(QtWidgets.QLineEdit, textboxName).text()) or len(self.findChild(QtWidgets.QLineEdit, textboxName).text()) == 0:
            self.findChild(QtWidgets.QLabel, headingName).setStyleSheet(HHHconf.design_textSmall)
            return True   
        else:
            self.findChild(QtWidgets.QLabel, headingName).setStyleSheet(HHHconf.design_textSmallError)
            return objName
    
    def validate_bankBSB(self, display=True):
        global old_format_bsb
        try:
            old_format_bsb
        except NameError:
            old_format_bsb = ''
        format_bsb = self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.bsb).text().replace('-', '')
        if len(format_bsb) > len(old_format_bsb):
            if len(format_bsb) < 3:
                output_bsb = format_bsb
            else:
                output_bsb = format_bsb[:3] + '-' + format_bsb[3:]
        else:
            output_bsb = self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.bsb).text()
        old_format_bsb = format_bsb
        self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.bsb).setText(output_bsb)

        if len(format_bsb) != 6 and len(format_bsb) > 0:
            self.findChild(QtWidgets.QLabel, self.headingPrefix + HHHconf.bsb).setStyleSheet(HHHconf.design_textSmallError)
            return HHHconf.bsb
        else:
            self.findChild(QtWidgets.QLabel, self.headingPrefix + HHHconf.bsb).setStyleSheet(HHHconf.design_textSmall)
            return True 

    def validate_bankACC(self, display=True):
        if len(self.findChild(QtWidgets.QLineEdit, self.editPrefix + HHHconf.acc).text()) == 0:
            # self.findChild(QtWidgets.QLabel, self.headingPrefix + HHHconf.acc).setStyleSheet(HHHconf.design_textSmallError)
            return HHHconf.acc
        else:
            self.findChild(QtWidgets.QLabel, self.headingPrefix + HHHconf.acc).setStyleSheet(HHHconf.design_textSmall)
            return True 

    def resetCursorPos(self):
        self.sender().setCursorPosition(0)
