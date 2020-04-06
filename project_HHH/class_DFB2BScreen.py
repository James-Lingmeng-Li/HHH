"""Back-to-Back Agreements

This screen is for the purposes of managing our Debtor Finanace back-to-back agreements

It is fairly incomplete, and should be expanded on when the business decides we need to. Currently, it contains the simple mechanism to request movements to the accounts team via email - but later should be made an inter-app communication (mPower)
"""

from PyQt5 import QtCore, QtGui, QtWidgets

import HHHconf, HHHfunc

class b2bScreen(HHHconf.HHHWindowWidget):

    signal_b2bUpdated = QtCore.pyqtSignal(object)

    def __init__(self, *args, **kwargs):
        super(b2bScreen, self).__init__(*args, **kwargs)
        self.populate_b2bScreen()

    def populate_b2bScreen(self):
        heading_b2b = QtWidgets.QLabel(self.title)
        heading_b2b.setStyleSheet(HHHconf.design_textMedium)
        heading_b2b.setMaximumHeight(20)
        self.vBox_contents.addWidget(heading_b2b)

        # Transactions Box
        self.transactionsBox = QtWidgets.QGroupBox()
        self.transactionsBox.setContentsMargins(5, 5, 5, 5)
        self.transactionsBox.setStyleSheet(HHHconf.design_GroupBox)
        self.vBox_contents.addWidget(self.transactionsBox)

        hBox_container = QtWidgets.QHBoxLayout()
        hBox_container.setContentsMargins(0, 0, 0, 0)
        hBox_container.setSpacing(0)
        hBox_container.setAlignment(QtCore.Qt.AlignTop)   
        self.transactionsBox.setLayout(hBox_container)

        # Transaction Input
        vBox_transactionDetails = QtWidgets.QVBoxLayout()
        vBox_transactionDetails.setContentsMargins(5, 5, 5, 5)
        vBox_transactionDetails.setSpacing(5)
        vBox_transactionDetails.setAlignment(QtCore.Qt.AlignTop)   
        hBox_container.addLayout(vBox_transactionDetails) 

        hBox_titleTransactions = QtWidgets.QHBoxLayout()
        hBox_titleTransactions.setContentsMargins(5, 5, 5, 5)
        hBox_titleTransactions.setSpacing(5)
        hBox_titleTransactions.setAlignment(QtCore.Qt.AlignTop)   
        vBox_transactionDetails.addLayout(hBox_titleTransactions) 

        heading_transactions = QtWidgets.QLabel('Add Transaction')
        heading_transactions.setStyleSheet(HHHconf.design_textMedium)
        heading_transactions.setMaximumHeight(20)
        hBox_titleTransactions.addWidget(heading_transactions)

        df_clients = HHHfunc.debtorFinanceEngine.getClientDetails(selectedClients=None, includeAQ=False, includeB2B=True)
        dict_clients = df_clients.set_index(HHHconf.DFClientInfo_AQClientName).to_dict('index')

        self.dict_transactionWidgets = {
            'selectClient':{
                'heading':'Select Client:', 
                'widget':HHHconf.HHHComboBox(parent=self.transactionsBox, placeholderText='None selected'), 
                'inputChanged':None, 
                'dataDict':dict_clients
            }, 
            'selectTransaction':{
                'heading':'Select Transaction:', 
                'widget':HHHconf.HHHComboBox(parent=self.transactionsBox, placeholderText='None selected'), 
                'inputChanged':self.validate_transaction, 
                'dataDict':HHHconf.dict_b2bTransactions
            }, 
            'amount':{
                'heading':'Amount:', 
                'widget':QtWidgets.QLineEdit(), 'placeholderText':'$', 
                'readOnly':False, 
                'inputChanged':self.validate_amount, 
                'validator':QtGui.QRegExpValidator(QtCore.QRegExp('^([0-9]\d*)(\.\d{0,2})?$'))
            }, 
            'gst':{
                'heading':'GST:', 
                'widget':QtWidgets.QLineEdit(), 
                'placeholderText':'$', 
                'readOnly':True, 
                'inputChanged':None, 
                'validator':QtGui.QRegExpValidator(QtCore.QRegExp('^([0-9]\d*)(\.\d{0,2})?$'))
            }, 
            'comments':{
                'heading':'Comments:', 
                'widget':HHHconf.HHHTextEditWidget(parent=self, placeholderText=''), 
                'placeholderText':'Comment required', 
                'inputChanged':None
            }
        }
        for obj, options in self.dict_transactionWidgets.items():
            hBox_current = QtWidgets.QHBoxLayout()
            hBox_current.setContentsMargins(15, 0, 0, 5)
            hBox_current.setSpacing(5)
            vBox_transactionDetails.addLayout(hBox_current) 

            currentHeading = QtWidgets.QLabel(options['heading'], objectName='heading_' + str(obj))
            currentHeading.setStyleSheet(HHHconf.design_textSmall)
            currentHeading.setFixedWidth(120)
            hBox_current.addWidget(currentHeading)

            currentWidget = options['widget']
            currentWidget.setObjectName('obj_' + str(obj))
            if isinstance(currentWidget, QtWidgets.QLineEdit):
                currentWidget.setStyleSheet(HHHconf.design_editBoxTwo)
                currentWidget.setPlaceholderText(options['placeholderText'])
                currentWidget.setReadOnly(options['readOnly'])
                if options['inputChanged']:
                    currentWidget.textChanged.connect(options['inputChanged'])
                if options['validator']:
                    currentWidget.setValidator(options['validator'])
            elif isinstance(currentWidget, HHHconf.HHHComboBox):
                if isinstance(options['dataDict'], dict):
                    for dataPoint in options['dataDict'].keys():
                        currentWidget.addItem(dataPoint)
                currentWidget.setCurrentIndex(-1)
                currentWidget.lineEdit().setReadOnly(True)
                if options['inputChanged']:
                    currentWidget.currentIndexChanged.connect(options['inputChanged'])
            elif isinstance(currentWidget, HHHconf.HHHTextEditWidget):
                currentWidget.setPlaceholderText(options['placeholderText'])
            hBox_current.addWidget(currentWidget)

        hBox_transactionButtons = QtWidgets.QHBoxLayout()
        hBox_transactionButtons.setContentsMargins(0, 0, 0, 0)
        hBox_transactionButtons.setSpacing(5)
        hBox_transactionButtons.setAlignment(QtCore.Qt.AlignTop)   
        vBox_transactionDetails.addLayout(hBox_transactionButtons)

        self.button_transactionClear = QtWidgets.QPushButton('Clear', clicked=self.clearTransaction, default=True)    
        self.button_transactionClear.setStyleSheet(HHHconf.design_smallButtonTransparent)
        self.button_transactionClear.setFixedHeight(20)
        hBox_transactionButtons.addWidget(self.button_transactionClear) 

        self.button_transactionCommit = QtWidgets.QPushButton('Commit Transaction', clicked=self.commitTransaction, default=True)
        self.button_transactionCommit.setStyleSheet(HHHconf.design_smallButtonTransparent)
        self.button_transactionCommit.setFixedHeight(20)
        hBox_transactionButtons.addWidget(self.button_transactionCommit)

        # # Pending transactions section
        # vBox_transactionsPending = QtWidgets.QVBoxLayout()
        # vBox_transactionsPending.setContentsMargins(15, 5, 5, 5)
        # vBox_transactionsPending.setSpacing(5)
        # vBox_transactionsPending.setAlignment(QtCore.Qt.AlignTop)   
        # hBox_container.addLayout(vBox_transactionsPending) 

        # hBox_titleTransactionsPending = QtWidgets.QHBoxLayout()
        # hBox_titleTransactionsPending.setContentsMargins(5, 5, 5, 5)
        # hBox_titleTransactionsPending.setSpacing(5)
        # hBox_titleTransactionsPending.setAlignment(QtCore.Qt.AlignTop)   
        # vBox_transactionsPending.addLayout(hBox_titleTransactionsPending) 

        # heading_transactionsPending = QtWidgets.QLabel('Pending Transactions')
        # heading_transactionsPending.setStyleSheet(HHHconf.design_textMedium)
        # heading_transactionsPending.setMaximumHeight(20)
        # hBox_titleTransactionsPending.addWidget(heading_transactionsPending)

        # self.transactionsPendingList = QtWidgets.QListWidget()
        # self.transactionsPendingList.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOn)
        # self.transactionsPendingList.setStyleSheet(HHHconf.design_listReadOnly + HHHconf.design_tableVerticalScrollBar.replace(HHHconf.tableHeaderColor, HHHconf.mainBackgroundColor))
        # self.transactionsPendingList.setMinimumWidth(180)
        # self.transactionsPendingList.setMinimumWidth(180)
        # self.transactionsPendingList.currentItemChanged.connect(lambda: print('test list'))
        # vBox_transactionsPending.addWidget(self.transactionsPendingList)

        # self.button_transactionPendingCancel = QtWidgets.QPushButton('Cancel Pending Transaction', clicked=(lambda: print('test cancel pending')), default=True)
        # self.button_transactionPendingCancel.setStyleSheet(HHHconf.design_smallButtonTransparent)
        # self.button_transactionPendingCancel.setFixedHeight(20)
        # vBox_transactionsPending.addWidget(self.button_transactionPendingCancel)

    def clearTransaction(self):
        for obj in self.dict_transactionWidgets.keys():
            widget = self.transactionsBox.findChild(QtWidgets.QWidget, name='obj_' + str(obj))
            if isinstance(widget, HHHconf.HHHComboBox):
                widget.setCurrentIndex(-1)
            elif isinstance(widget, QtWidgets.QLineEdit):
                widget.setText('')
            elif isinstance(widget, HHHconf.HHHTextEditWidget):
                widget.clear()
    
    def commitTransaction(self):
        HHHconf.widgetEffects.flashMessage(label=self.heading_status, duration=3000, message='Committing transaction. Please wait...', messageType='warning')
        # Check if all inputs are valid
        invalidInputs = []
        for obj in self.dict_transactionWidgets.keys():
            widget = self.transactionsBox.findChild(QtWidgets.QWidget, name='obj_' + str(obj))
            if isinstance(widget, HHHconf.HHHComboBox):
                if widget.currentIndex() == -1:
                    invalidInputs.append(self.dict_transactionWidgets[obj]['heading'].replace(':', ''))
            elif isinstance(widget, QtWidgets.QLineEdit):
                if len(widget.text()) == 0:
                    invalidInputs.append(self.dict_transactionWidgets[obj]['heading'].replace(':', ''))
            elif isinstance(widget, HHHconf.HHHTextEditWidget):
                if len(widget.toPlainText()) == 0:
                    invalidInputs.append(self.dict_transactionWidgets[obj]['heading'].replace(':', ''))
        if invalidInputs:
            return HHHconf.widgetEffects.flashMessage(label=self.heading_status, duration=3000, message='Invalid: <html style="font-weight:bold;">' + ', '.join(invalidInputs) + '</html>. Transaction not commited', messageType='error')

        # Commit Transaction (Temporary = send email to accounts team, TODO = write transaction to mPower instruction tables)
        emailBody = ('''
        <div>
            Hi Accounts, <br><br>Please perform the following back-to-back transaction:<br><br>
            <div style=\"margin-left:4em;font-family:Courier\">
                Client: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;''' + self.transactionsBox.findChild(HHHconf.HHHComboBox, name='obj_selectClient').currentText() + '''<br>
                Transaction: &nbsp;''' + self.transactionsBox.findChild(HHHconf.HHHComboBox, name='obj_selectTransaction').currentText() + '''<br>
                Amount: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;''' + HHHconf.moneyFormat.format(float(self.transactionsBox.findChild(QtWidgets.QLineEdit, name='obj_amount').text())) + '''<br>
                GST: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;''' + HHHconf.moneyFormat.format(float(self.transactionsBox.findChild(QtWidgets.QLineEdit, name='obj_gst').text())) + '''
            </div>
            <br><br>Comments:<br><br>
            <div style=\"margin-left:4em;font-family:Courier\">
                \"''' + self.transactionsBox.findChild(HHHconf.HHHTextEditWidget, name='obj_comments').toPlainText() + '''\" - ''' + self.username + '''
            </div>
        </div>
        ''')
        HHHfunc.mailMergeEngine.sendEmail(outlookSession=HHHfunc.mailMergeEngine.newOutlookSession(), sender=HHHfunc.mailMergeEngine.PrimarySmtpAddress(), recipient={'to':HHHconf.email_accountsTeam, 'cc':None, 'bcc':None}, subject='B2B Transaction Request', body=emailBody, attachments=[])
        self.signal_b2bUpdated.emit({HHHconf.events_EventTime: None, HHHconf.events_UserName: HHHconf.username_PC, HHHconf.events_Event: HHHconf.dict_eventCategories['b2b'], HHHconf.events_EventDescription: 'Requested ' + HHHconf.moneyFormat.format(float(self.transactionsBox.findChild(QtWidgets.QLineEdit, name='obj_amount').text()))  + ' ' + str(self.transactionsBox.findChild(HHHconf.HHHComboBox, name='obj_selectTransaction').currentText()) + ' for ' + str(self.transactionsBox.findChild(HHHconf.HHHComboBox, name='obj_selectClient').currentText())})
        self.clearTransaction()        
        return HHHconf.widgetEffects.flashMessage(label=self.heading_status, duration=3000, message='Transaction committed successfully. Email the Accounts team for any reversals or issues', messageType='success')

    def validate_transaction(self):
        selectedTransaction = self.transactionsBox.findChild(HHHconf.HHHComboBox, name='obj_selectTransaction').currentText()
        if selectedTransaction == 'Same Day Payment':
            self.transactionsBox.findChild(HHHconf.HHHTextEditWidget, name='obj_comments').setPlainText('Please charge them a Same Day Payment Fee in accordance with their Fee Schedule.')
        else:
            self.transactionsBox.findChild(HHHconf.HHHTextEditWidget, name='obj_comments').clear()
        self.validate_amount()

    def validate_amount(self):
        currentAmount = float(self.transactionsBox.findChild(QtWidgets.QLineEdit, name='obj_amount').text()) if len(self.transactionsBox.findChild(QtWidgets.QLineEdit, name='obj_amount').text()) > 0 else None
        if currentAmount is None:
            return self.transactionsBox.findChild(QtWidgets.QLineEdit, name='obj_gst').setText('')
        selectedTransaction = self.transactionsBox.findChild(HHHconf.HHHComboBox, name='obj_selectTransaction').currentText()
        if len(selectedTransaction) > 0:
            if HHHconf.dict_b2bTransactions[selectedTransaction]['includeGST'] is True:
                return self.transactionsBox.findChild(QtWidgets.QLineEdit, name='obj_gst').setText('{:0.2f}'.format(currentAmount * HHHconf.gstComponent))
        self.transactionsBox.findChild(QtWidgets.QLineEdit, name='obj_gst').setText(str('0.00'))