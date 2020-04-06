"""Main App

This is where the code initializes for the HHH_plugin in mPower

The entire code requires the following modules to be installed alongside the Python environment you are running:
    - Pandas: for large data manipulation
    - PyQt5: for GUI creation
    - Selenium: for web-automation
    - sqlalchemy: for connecting with the SQL server
    - argon2: for one-way password encryptions
    - Fernet: for two-way password encryption/decryption`
    - docx: for Microsoft Word documents
    - openpyxl: for Microsoft Excel documents
    - dateutil: for working with datetime

If you are missing a module namespace, please try a 'pip install' in the terminal or try downloading it as it may not be a standary library module.

It is recommended to setup a Python virtual environment for this project

"""

import datetime
print('Importing resources...')
start = datetime.datetime.now()

import sys, os, math, subprocess, pandas, datetime, time, calendar, numpy, pyodbc, re, openpyxl, docx, shutil, threading, traceback, ctypes

from PyQt5 import QtCore, QtGui, QtWidgets

import HHHconf, HHHfunc, class_EFAddAgreementScreen, class_EFViewAgreementScreen, class_EFAmendPaymentScreen, class_EFViewClientScreen, class_settingsScreen, class_mailMergerScreen, class_DFBankingScreen, class_DFB2BScreen

# DPI scaling handling
errorCode = ctypes.windll.shcore.SetProcessDpiAwareness(2)
if hasattr(QtCore.Qt, 'AA_EnableHighDpiScaling'):
    QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling, True)
if hasattr(QtCore.Qt, 'AA_UseHighDpiPixmaps'):
    QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_UseHighDpiPixmaps, True)

# Main Screen
class MainApp(QtWidgets.QWidget):

    def __init__(self,username=None):
        super().__init__()
        if username is not None:
            self.username = username
            self.loginTime = datetime.datetime.now()
        self.subMenu_paymentStatusFilter = []    
        self.list_closedAgreements = ['EF0000000'] # default value required otherwise if no closed agreements, EFPaymentsTable loads blank
        self.loadHHH()
        # necessary, as 3rd party windows may slide in between main window and children windows, have not solved for this issue otherwise 
        self.setWindowFlags(self.windowFlags() | QtCore.Qt.WindowStaysOnTopHint | QtCore.Qt.X11BypassWindowManagerHint)
        self.setWindowState(self.windowState() & ~QtCore.Qt.WindowMinimized | QtCore.Qt.WindowActive)
        self.show()
        HHHfunc.mainEngine.updateLoginDetails(username=self.username,data={HHHconf.userAdmin_latestLogin: self.loginTime})
  
    def showEvent(self,event):
        try:
            self.childWindow.setFocus()
        except:
            pass

    def closeEvent(self,event):
        try:
            self.childWindow.close()
        except:
            pass

    def moveEvent(self,event):
        try:
            self.childWindow.move(self.geometry().x(),self.geometry().y())
        except:
            pass

    def loadHHH(self):
        self.setWindowTitle(HHHconf.name_app)
        self.setGeometry(HHHconf.appLeft,HHHconf.appTop,HHHconf.appWidth,HHHconf.appHeight)

        # Background
        vBox_background = QtWidgets.QVBoxLayout(self)
        vBox_background.setContentsMargins(0, 0, 0, 0)
        background = QtWidgets.QLabel()
        pixmapBackground = QtGui.QPixmap(HHHconf.icons_dir + '\\background_galaxy.jpg')
        print(pixmapBackground.devicePixelRatio())
        background.setPixmap(pixmapBackground)
        self.resize(pixmapBackground.width(), pixmapBackground.height())
        self.setMinimumSize(pixmapBackground.width(), pixmapBackground.height())
        self.setMaximumSize(pixmapBackground.width(), pixmapBackground.height())
        vBox_background.addWidget(background)

        vBox_contents = QtWidgets.QVBoxLayout()
        vBox_contents.setContentsMargins(0, 0, 0, 0)
        vBox_contents.setSpacing(0)
        background.setLayout(vBox_contents)

        hBox_topButtons = QtWidgets.QHBoxLayout()
        hBox_topButtons.setContentsMargins(0, 0, 0, 0)
        hBox_topButtons.setSpacing(0)
        vBox_contents.addLayout(hBox_topButtons)

        hBox_topButtons.addStretch(1)   

        button_settings = QtWidgets.QPushButton('Settings')
        button_settings.setFocusPolicy(QtCore.Qt.NoFocus)
        button_settings.setStyleSheet(HHHconf.design_textButton)
        button_settings.setMinimumHeight(25)
        button_settings.clicked.connect(self.open_settingsScreen)
        hBox_topButtons.addWidget(button_settings)  

        # Tabs
        HHH_tabs = QtWidgets.QTabWidget()
        tabDebtor = QtWidgets.QWidget()
        tabEquipment = QtWidgets.QWidget() 
        HHH_tabs.addTab(tabDebtor, HHHconf.name_tabDebtor)
        HHH_tabs.addTab(tabEquipment, HHHconf.name_tabEquipment)

        vBox_contents.addWidget(HHH_tabs) 
        HHH_tabs.setStyleSheet(HHHconf.design_tabs)
        HHH_tabs.setFocusPolicy(QtCore.Qt.NoFocus)

        # Status List - at the bottom
        self.statusList = QtWidgets.QListWidget()
        self.statusList.setFocusPolicy(QtCore.Qt.NoFocus)
        self.statusList.setMinimumSize(QtCore.QSize(pixmapBackground.width(), 0))
        self.statusList.setMaximumSize(QtCore.QSize(pixmapBackground.width(), 100))
        self.statusList.setStyleSheet(HHHconf.design_listReadOnly)
        self.statusList.setWordWrap(True)
        self.update_statusList('Logged in as ' + self.username)
        vBox_contents.addWidget(self.statusList)          
        # Tab Contents 
        # Debtor Finance Tab
        self.vBox_debtor = QtWidgets.QVBoxLayout(tabDebtor) 
        self.vBox_debtor.setContentsMargins(10, 10, 10, 10)
        self.vBox_debtor.setSpacing(5)
        self.vBox_debtor.setAlignment(QtCore.Qt.AlignTop)

        self.button_openBankingScreen = QtWidgets.QPushButton('Banking', clicked=self.open_bankingScreen, default=True)    
        self.button_openBankingScreen.setStyleSheet(HHHconf.design_smallButtonTransparent)
        self.button_openBankingScreen.setFixedHeight(20)
        self.vBox_debtor.addWidget(self.button_openBankingScreen)

        self.button_open_b2bScreen = QtWidgets.QPushButton('Back to Back Agreements', clicked=self.open_b2bScreen, default=True)    
        self.button_open_b2bScreen.setStyleSheet(HHHconf.design_smallButtonTransparent)
        self.button_open_b2bScreen.setFixedHeight(20)
        self.vBox_debtor.addWidget(self.button_open_b2bScreen)

        self.button_mailMerger = QtWidgets.QPushButton('Mail Merger', clicked=self.open_mailMergerScreen, default=True)    
        self.button_mailMerger.setStyleSheet(HHHconf.design_smallButtonTransparent)
        self.button_mailMerger.setFixedHeight(20)
        self.vBox_debtor.addWidget(self.button_mailMerger)

        self.button_halo = QtWidgets.QPushButton('Halo', clicked=(lambda: self.beginHaloThread(application='HALO')), default=True)    
        self.button_halo.setStyleSheet(HHHconf.design_smallButtonTransparent)
        self.button_halo.setFixedHeight(20)
        self.vBox_debtor.addWidget(self.button_halo)

        self.button_haloDemo = QtWidgets.QPushButton('Halo Demo', clicked=(lambda: self.beginHaloThread(application='HALO_DEMO')), default=True)    
        self.button_haloDemo.setStyleSheet(HHHconf.design_smallButtonTransparent)
        self.button_haloDemo.setFixedHeight(20)
        self.vBox_debtor.addWidget(self.button_haloDemo)

        # Equipment Tab
        self.vBox_Equipment = QtWidgets.QVBoxLayout(tabEquipment)
        self.vBox_Equipment.setContentsMargins(10, 10, 10, 10)
        self.vBox_Equipment.setSpacing(5)
        self.vBox_Equipment.setAlignment(QtCore.Qt.AlignTop) 

        self.hBox_EFAgreementsSelect = QtWidgets.QHBoxLayout()
        self.hBox_EFAgreementsSelect.setContentsMargins(0, 0, 0, 0)
        self.hBox_EFAgreementsSelect.setSpacing(5)
        self.hBox_EFAgreementsSelect.setAlignment(QtCore.Qt.AlignTop)   
        self.vBox_Equipment.addLayout(self.hBox_EFAgreementsSelect)

        self.button_EFAddAgreement = QtWidgets.QPushButton('Add Agreement', clicked=self.EFAddAgreementScreen, default=True)
        self.button_EFAddAgreement.setMinimumSize(QtCore.QSize(50,50))
        self.button_EFAddAgreement.setStyleSheet(HHHconf.design_largeButtonTransparent)
        self.hBox_EFAgreementsSelect.addWidget(self.button_EFAddAgreement) 

        self.EFAgreementsTable = HHHconf.HHHTableWidget(defaultSectionSize=130, fixedHeight=330, cornerWidget='copy',contextMenu=self.EFAgreementsTableMenu,doubleClicked=self.EFAgreementsTable_doubleClicked)
        self.vBox_Equipment.addWidget(self.EFAgreementsTable)

        self.EFPaymentsTable = HHHconf.HHHTableWidget(defaultSectionSize=130, fixedHeight=300, cornerWidget='copy',contextMenu=self.EFPaymentsTableMenu,doubleClicked=self.EFPaymentsTable_doubleClicked)
        self.vBox_Equipment.addWidget(self.EFPaymentsTable)

        # TEMPORARY - Final Buttons EF TESTS TODO - make BOD and EOD automatic and scheduled (could be based on events)
        hBox_buttons = QtWidgets.QHBoxLayout()
        hBox_buttons.setContentsMargins(0, 0, 0, 0)
        hBox_buttons.setSpacing(10)
        hBox_buttons.setAlignment(QtCore.Qt.AlignTop)  
        self.vBox_Equipment.addLayout(hBox_buttons)

        self.button_testBOD = QtWidgets.QPushButton('Test Beginning of Day', clicked=HHHfunc.equipmentFinanceEngine.startOfDay, default=True)
        self.button_testBOD.setFixedHeight(20)
        self.button_testBOD.setStyleSheet(HHHconf.design_smallButtonTransparent)
        hBox_buttons.addWidget(self.button_testBOD) 

        self.button_testEOD = QtWidgets.QPushButton('Test End of Day', clicked=HHHfunc.equipmentFinanceEngine.endOfDay, default=True)
        self.button_testEOD.clicked.connect(lambda: self.update_EFAgreementsTable())
        self.button_testEOD.clicked.connect(lambda: self.update_EFPaymentsTable())
        self.button_testEOD.setFixedHeight(20)
        self.button_testEOD.setStyleSheet(HHHconf.design_smallButtonTransparent)
        hBox_buttons.addWidget(self.button_testEOD) 
        self.update_EFAgreementsTable()                
        self.update_EFPaymentsTable(filterList=HHHconf.defaultEFPaymentsTableMenuFilter)        

    # Functions
    def open_settingsScreen(self):
        if not HHHfunc.mainEngine.requestPassword(self, self.username):
            return
        self.childWindow =  class_settingsScreen.settingsScreen(username=self.username, geometry=self.geometry(), padding={'left':0, 'top':0, 'right':0, 'bottom':self.geometry().height() / 2}, title='Settings', updateButton='Update Email Address', parent=self)
        self.childWindow.show()
        self.childWindow.activateWindow()
        self.childWindow.setFocus()
        self.childWindow.signal_settingsUpdated.connect(lambda x: self.update_statusList(contents=x, display=False))
    
    def update_statusList(self, contents, display=True):
        # contents: Pass a string for display-only, pass a dictionary (HHHconf.dict_toSQL_Events) to add to audit trail db (and set display True or False for display)
        if isinstance(contents, str):
            self.statusList.addItem(str(datetime.datetime.now().strftime(HHHconf.dateTimeFormat)) + ': ' + contents)
        if isinstance(contents, dict):
            newEventID = HHHfunc.mainEngine.addEvent(contents)
            if display is True:
                self.statusList.addItem(str(contents[HHHconf.events_EventTime].strftime(HHHconf.dateTimeFormat) + ': ID(' + newEventID + ') ' + contents[HHHconf.events_Event] + ' - ' + contents[HHHconf.events_EventDescription]))
        self.statusList.scrollToBottom()

    # Debtor Finance Tab Functions
    def open_bankingScreen(self):
        self.childWindow =  class_DFBankingScreen.bankingScreen(username=self.username, geometry=self.geometry(), padding={'left':0, 'top':20, 'right':0, 'bottom':20}, title='Banking', updateButton=None, parent=self)
        self.childWindow.show()
        self.childWindow.activateWindow()
        self.childWindow.setFocus()
        self.childWindow.signal_animateWindow.connect(lambda x: HHHconf.widgetEffects.translucifyWindow(self, x, self.geometry()))
        self.childWindow.signal_animateWindow.connect(lambda x: HHHconf.widgetEffects.animateWindowOffScreen(self, x, self.geometry()))
        self.childWindow.signal_addEvent.connect(lambda x: self.update_statusList(contents=x, display=True))
    
    def open_b2bScreen(self):  
        self.childWindow = class_DFB2BScreen.b2bScreen(username=self.username, geometry=self.geometry(), title='Back to Back Agreeements', parent=self)
        self.childWindow.show()
        self.childWindow.activateWindow()
        self.childWindow.setFocus()
        self.childWindow.signal_b2bUpdated.connect(lambda x: self.update_statusList(contents=x, display=True))

    def open_mailMergerScreen(self):
        self.childWindow = class_mailMergerScreen.mailMergerScreen(username=self.username, geometry=self.geometry(), padding={'left':0, 'top':20, 'right':0, 'bottom':20}, title='Mail Merger', updateButton=None, parent=self)
        self.childWindow.show()
        self.childWindow.activateWindow()
        self.childWindow.setFocus()
        # TODO: proper status list updates for non-test Mail Merges only, as well as saving results of merge threads to the events db
        # self.childWindow.signal_mailMergeStarted.connect(lambda: self.update_statusList(contents='Initiated a Mail Merge', display=True))
        # self.childWindow.signal_mailMergeComplete.connect(lambda: self.update_statusList(contents='Completed a Mail Merge', display=True))
        

    def beginHaloThread(self, application:str='HALO_DEMO'):
        # TODO: Clean up code - repetitive code in this method for both HALO and HALO_DEMO
        # Get app credentials 
        df_appCredentials = HHHfunc.mainEngine.getAppCredentials(username=self.username)
        try:
            dict_haloCredentials = df_appCredentials[df_appCredentials[HHHconf.userCredentials_application] == application].to_dict('records')[0]
        except IndexError:
            return HHHconf.widgetEffects.flashMessage(self.heading_receiptsMessage, duration=3000, message='Halo credentials not found. Contact the admin', messageType='error')
        # Password confirm screen
        if not HHHfunc.mainEngine.requestPassword(self, self.username, title='Confirm Action', dialogText='You are about to log in to ' + application + '. Please enter your password to continue:'):
            return
        # Initiate receipts thread (so GUI does not time out)
        if application == 'HALO':
            try:
                if self.haloThread.isRunning():
                    return self.update_statusList(str(application) + ' worker already at work. Please wait 20 seconds before trying again.')
            except:
                pass
            self.haloThread = QtCore.QThread()
            self.haloWorker = HHHfunc.debtorFinanceEngine.haloWorker(app=application, parent=self, username=str(dict_haloCredentials[HHHconf.userCredentials_username]), password=HHHfunc.mainEngine.twoWayDecrypt(str(dict_haloCredentials[HHHconf.userCredentials_password])))
            self.haloWorker.moveToThread(self.haloThread)
            self.haloThread.started.connect(self.haloWorker.haloLogin)
            self.haloWorker.signal_finished.connect(lambda x: self.update_statusList(contents=x, display=True)) # audit required only for non-demo login
            self.haloWorker.signal_finished.connect(self.haloThread.quit)            
            self.haloWorker.signal_finished.connect(self.haloWorker.deleteLater)
            self.haloWorker.signal_passwordUpdated.connect(lambda x: HHHfunc.mainEngine.updateAppCredentials(app=application, data={HHHconf.userCredentials_password: HHHfunc.mainEngine.twoWayEncrypt(x)}, username=self.username))
            self.haloThread.start()
        elif application == 'HALO_DEMO':
            try:
                if self.haloDemoThread.isRunning():
                    return self.update_statusList(str(application) + ' worker already at work. Please wait 20 seconds before trying again.')
            except:
                pass
            self.haloDemoThread = QtCore.QThread()
            self.haloDemoWorker = HHHfunc.debtorFinanceEngine.haloWorker(app=application, parent=self, username=str(dict_haloCredentials[HHHconf.userCredentials_username]), password=HHHfunc.mainEngine.twoWayDecrypt(str(dict_haloCredentials[HHHconf.userCredentials_password])))
            self.haloDemoWorker.moveToThread(self.haloDemoThread)
            self.haloDemoThread.started.connect(self.haloDemoWorker.haloLogin)
            self.haloDemoWorker.signal_finished.connect(self.haloDemoThread.quit)
            self.haloDemoWorker.signal_finished.connect(self.haloDemoWorker.deleteLater)
            self.haloDemoWorker.signal_passwordUpdated.connect(lambda x: HHHfunc.mainEngine.updateAppCredentials(app=application, data={HHHconf.userCredentials_password: HHHfunc.mainEngine.twoWayEncrypt(x)}, username=self.username))
            self.haloDemoThread.start()

    # Equiment Finance Functions
    def EFAddAgreementScreen(self):
        self.childWindow = class_EFAddAgreementScreen.EFAddAgreementScreen(username=self.username, geometry=self.geometry(),padding={'left':0,'top':20,'right':0,'bottom':20}, title='Equipment Finance - Add Agreement', updateButton='Activate Client', parent=self)
        self.childWindow.show()
        self.childWindow.activateWindow()
        self.childWindow.setFocus()
        self.childWindow.signal_update_EFAgreementsTable.connect(lambda: self.update_EFAgreementsTable())
        self.childWindow.signal_update_EFAgreementsTable.connect(lambda: self.update_EFPaymentsTable())
        self.childWindow.signal_agreementAdded.connect(lambda x: self.update_statusList(contents=x, display=True))

    def EFAgreementsTableMenu(self, pos):
        #TODO repurpose the context menu if required - or delete if not
        # if self.EFAgreementsTable.currentIndex() is None:
        #     return
        # selectedAgreement = str(self.EFAgreementsTable.model().index(self.EFAgreementsTable.currentIndex().row(), 0).data(QtCore.Qt.DisplayRole))

        # menu = QtWidgets.QMenu()
        # menu.setStyleSheet(HHHconf.design_tableMenu)
        # Action_openClient = menu.addAction('Open Agreement')
        # action = menu.exec_(QtGui.QCursor.pos())
        # if action == Action_openClient:
        #     self.EFViewAgreementScreen(selectedAgreement)
        pass

    def EFAgreementsTable_doubleClicked(self,clickedLine):
        selectedAgreement = str(self.EFAgreementsTable.model().index(self.EFAgreementsTable.currentIndex().row(), 0).data(QtCore.Qt.DisplayRole))
        self.EFViewAgreementScreen(selectedAgreement)

    def EFViewAgreementScreen(self,selectedAgreement):   
        self.childWindow =  class_EFViewAgreementScreen.EFViewAgreementScreen(selectedAgreement=selectedAgreement, username=self.username, geometry=self.geometry(),padding={'left':0,'top':20,'right':0,'bottom':20}, title='Equipment Finance - View Agreement',updateButton='Update Agreement Notes', parent=self)
        self.childWindow.show()
        self.childWindow.activateWindow()
        self.childWindow.setFocus()
        self.childWindow.signal_agreementUpdated.connect(lambda x: self.update_EFAgreementsTable())
        self.childWindow.signal_agreementUpdated.connect(lambda x: self.update_EFPaymentsTable())
        self.childWindow.signal_agreementUpdated.connect(lambda x: self.update_statusList(contents=x, display=True))

    def update_EFAgreementsTable(self):
        #Get Data from db
        self.df_allClients = HHHfunc.equipmentFinanceEngine.getClientDetails()
        self.df_allAgreements = HHHfunc.equipmentFinanceEngine.getAgreementDetails()
        self.df_allSubordinates = HHHfunc.equipmentFinanceEngine.getCurrentSubordinates()

        # Configure new DF to use for display
        df_displayAgreements = pandas.merge(left=self.df_allAgreements, right=self.df_allClients, how='left', left_on=HHHconf.client_number, right_on=HHHconf.client_number)
        df_displayAgreements = df_displayAgreements[[HHHconf.agreement_number, HHHconf.client_name, HHHconf.original_balance, HHHconf.periodic_repayment]]
        # Drop closed agreements
        self.list_closedAgreements += self.df_allAgreements[self.df_allAgreements[HHHconf.agreement_status] == HHHconf.dict_agreementStatus[HHHconf.agreementStatus_closed]][HHHconf.agreement_number].tolist()
        df_displayAgreements = df_displayAgreements[~df_displayAgreements[HHHconf.agreement_number].isin(self.list_closedAgreements)]
        # Add columns that aren't stored in db
        df_displayAgreements['Outstanding'] = [HHHconf.moneyFormat.format(self.df_allSubordinates[(self.df_allSubordinates[HHHconf.payment_number].str.contains(x)) & (self.df_allSubordinates[HHHconf.subordinate_status].isin(HHHconf.outstandingStatuses))][HHHconf.subordinate_amount].sum()) for x in df_displayAgreements[HHHconf.agreement_number]]
        df_displayAgreements['Overdue'] = [HHHconf.moneyFormat.format(self.df_allSubordinates[(self.df_allSubordinates[HHHconf.payment_number].str.contains(x)) & (self.df_allSubordinates[HHHconf.subordinate_status] == HHHconf.dict_paymentStatus[HHHconf.paymentStatus_overdue])][HHHconf.subordinate_amount].sum()) for x in df_displayAgreements[HHHconf.agreement_number]]
        df_displayAgreements['Next Payment'] = [self.df_allSubordinates[(self.df_allSubordinates[HHHconf.payment_number].str.contains(x)) & (self.df_allSubordinates[HHHconf.subordinate_status] == HHHconf.dict_paymentStatus[HHHconf.paymentStatus_pending])][HHHconf.subordinate_value_date].min().strftime(HHHconf.dateFormat2) for x in df_displayAgreements[HHHconf.agreement_number]]

        # Format columns
        df_displayAgreements[HHHconf.original_balance] = df_displayAgreements[HHHconf.original_balance].map(HHHconf.moneyFormat.format)
        df_displayAgreements[HHHconf.periodic_repayment] = df_displayAgreements[HHHconf.periodic_repayment].map(HHHconf.moneyFormat.format)
        df_displayAgreements.rename({HHHconf.agreement_number:'Agreement', HHHconf.client_name:'Client', HHHconf.original_balance:'Original Balance', HHHconf.periodic_repayment:'Periodic Repayment'}, axis='columns', inplace=True)

        self.EFAgreementsTable.setModel(df_displayAgreements)

    def update_EFPaymentsTable(self, filterList:list=None):
        self.EFPaymentsTable.setFocus(True)
        df_currentSubordinates = HHHfunc.equipmentFinanceEngine.getCurrentSubordinates(collapseGST=True)
        # Filter List based on menu selection
        if filterList is None:
            try:
                filterList = self.currentEFPaymentsTableMenuFilter
            except:
                filterList = HHHconf.defaultEFPaymentsTableMenuFilter

        df_currentSubordinates = df_currentSubordinates[df_currentSubordinates[HHHconf.subordinate_status].isin(filterList)]
        self.currentEFPaymentsTableMenuFilter = filterList
        
        df_currentSubordinates = df_currentSubordinates[[HHHconf.subordinate_value_date, HHHconf.subordinate_number, HHHconf.subordinate_type, HHHconf.subordinate_amount, HHHconf.subordinate_status]]
        
        # Display in Table
        df_currentSubordinates.insert(loc=2, column=HHHconf.client_name, value=0)
        df_currentSubordinates.sort_values(by=[HHHconf.subordinate_value_date, HHHconf.subordinate_number], na_position='first', inplace=True)
        df_currentSubordinates.reset_index(drop=True, inplace=True) # must re-index df
        
        # drop closed agreement transactions
        df_currentSubordinates = df_currentSubordinates[~df_currentSubordinates[HHHconf.subordinate_number].str.contains('|'.join(self.list_closedAgreements)) ]
        df_temp = pandas.merge(left=self.df_allAgreements, right=self.df_allClients, how='left', left_on=HHHconf.client_number, right_on=HHHconf.client_number)[[HHHconf.agreement_number, HHHconf.client_name]]
        df_currentSubordinates[HHHconf.client_name] = [df_temp[df_temp[HHHconf.agreement_number] == x.split('_')[0]][HHHconf.client_name].iloc[0] for x in df_currentSubordinates[HHHconf.subordinate_number]]
        df_currentSubordinates[HHHconf.subordinate_value_date] = df_currentSubordinates[HHHconf.subordinate_value_date].dt.strftime(HHHconf.dateFormat2)
        df_currentSubordinates[HHHconf.subordinate_amount] = df_currentSubordinates[HHHconf.subordinate_amount].map(HHHconf.moneyFormat.format)

        df_currentSubordinates.rename({HHHconf.subordinate_value_date:'Value Date', HHHconf.client_name:'Client', HHHconf.subordinate_number:'Subordinate Number', HHHconf.subordinate_type:'Subordinate Type', HHHconf.subordinate_amount:'Amount', HHHconf.subordinate_status:'Status'}, axis='columns', inplace=True)
        self.EFPaymentsTable.setModel(df_currentSubordinates)

    def EFPaymentsTableMenu(self, pos):
        if self.EFPaymentsTable.currentIndex() is None:
            return
        selectedSubordinate = str(self.EFPaymentsTable.model().index(self.EFPaymentsTable.currentIndex().row(), 1).data(QtCore.Qt.DisplayRole))

        if not self.subMenu_paymentStatusFilter: # Create default subMenu on generation of menu
            for key in HHHconf.dict_paymentStatus.keys():
                current_paymentStatusMenuItem = {}
                current_paymentStatusMenuItem['type'] = 'action'
                current_paymentStatusMenuItem['displayText'] = HHHconf.dict_paymentStatus[key]
                current_paymentStatusMenuItem['checkable'] = True
                # Default filters applied
                if current_paymentStatusMenuItem['displayText'] in HHHconf.defaultEFPaymentsTableMenuFilter:
                    current_paymentStatusMenuItem['checkedState'] = True
                else:
                    current_paymentStatusMenuItem['checkedState'] = False
                current_paymentStatusMenuItem['connect'] = 'test'
                self.subMenu_paymentStatusFilter.append(current_paymentStatusMenuItem)

        menuItems = [
                {'type':'action', 'displayText': 'Open Payment', 'checkable': False, 'checkedState': False, 'connect':'test'},
                {'type':'menu', 'displayText': 'Status Filter', 'subMenu': self.subMenu_paymentStatusFilter}
            ]

        EFPaymentsTable_Menu = HHHconf.CustomMenu(menuItems, parent=self.EFPaymentsTable)
        EFPaymentsTable_Menu.signal_checkableAction.connect(self.filter_EFPaymentsTable)
        EFPaymentsTable_Menu.menuAction = EFPaymentsTable_Menu.exec_(QtGui.QCursor.pos())

        for item in menuItems:
            if 'object' in item:
                if EFPaymentsTable_Menu.menuAction == item['object']:
                    if item['displayText'] == 'Open Payment' and selectedSubordinate:    
                        self.EFPaymentsTable_doubleClicked(selectedSubordinate)
                    return
        
    def filter_EFPaymentsTable(self,trigger):
        visibleStatusList = []
        for item in self.subMenu_paymentStatusFilter:
            if 'object' in item:
                if trigger == item['object']:
                    item['checkedState'] = not item['checkedState']
                if item['checkedState'] == True:
                    visibleStatusList.append(item['displayText'])
        self.update_EFPaymentsTable(filterList=visibleStatusList)

    def EFPaymentsTable_doubleClicked(self,clickedLine):
        selectedSubordinate = str(self.EFPaymentsTable.model().index(self.EFPaymentsTable.currentIndex().row(), 1).data(QtCore.Qt.DisplayRole))
        selectedPayment = selectedSubordinate.split('_')[0] + '_' + selectedSubordinate.split('_')[1]
        self.open_EFAmendPaymentScreen(selectedPayment)

    def open_EFAmendPaymentScreen(self,selectedPayment):
        selectedAgreement = selectedPayment.split('_')[0]
        self.childWindow =  class_EFAmendPaymentScreen.EFAmendPaymentScreen(selectedAgreement=selectedAgreement, selectedPayment=selectedPayment, username=self.username, geometry=self.geometry(), padding={'left':0, 'top':20, 'right':0, 'bottom':220}, title='Amend Payment Details', updateButton='Update Payment Notes', parent=self)
        self.childWindow.show()
        self.childWindow.activateWindow()
        self.childWindow.setFocus()
        self.childWindow.signal_paymentUpdated.connect(lambda x: self.update_EFAgreementsTable())
        self.childWindow.signal_paymentUpdated.connect(lambda x: self.update_EFPaymentsTable())
        self.childWindow.signal_paymentUpdated.connect(lambda x: self.update_statusList(contents=x, display=True))

print('Initialising application...')

# Execution
if __name__=='__main__':
    try:
        # Check if Moneytech common directory found. If not, then computer is not connected to moneytech servers and cannot use HHH
        if not os.path.isdir(HHHconf.common_dir):
            msgbox = ctypes.windll.user32.MessageBoxA
            msgbox(0, 'Moneytech connection not detected. Closing...', 'ERROR', 0)
        else:
            app = QtWidgets.QApplication(sys.argv)

            # TEST DPIs - TODO delete for deploy
            for screen in app.screens():
                print(str(screen.physicalDotsPerInch()))

            ex = MainApp(username=HHHconf.username_PC)
            print('Load time ' + str(datetime.datetime.now() - start))
            print('Welcome ' + HHHconf.username_PC)
            sys.exit(app.exec_())
    except Exception as e:
        print('ERROR: ' + str(e) + ' ' + str(traceback.format_exc()))
        HHHfunc.mainEngine.addEvent(dict_event={HHHconf.events_EventTime: datetime.datetime.now(), HHHconf.events_UserName: HHHconf.username_PC, HHHconf.events_Event: HHHconf.dict_eventCategories['error'], HHHconf.events_EventDescription: str(e)})
        