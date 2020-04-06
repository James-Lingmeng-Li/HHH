import sys, os, pandas, re, docx, shutil, threading, traceback, copy

from PyQt5 import QtCore, QtGui, QtWidgets

import HHHconf, HHHfunc

class mailMergerScreen(HHHconf.HHHWindowWidget):

    signal_mailMergeStarted = QtCore.pyqtSignal()
    signal_mailMergeComplete = QtCore.pyqtSignal()

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.tempMergeDir = HHHconf.temp_dir + '\\_uploaded\\'
        self.defaultNumberOfTestEmails = HHHconf.defaultNumberofTestEmails
        self.numberOfTestEmails = self.defaultNumberOfTestEmails
        self.dict_uploadedDocs = {} # All uploaded docs go here
        self.dict_foundVars = {} # All vars found from uploaded docs go here  
        self.dict_uploadedData = {
            'path':None, 
            'dataframe':None
            } # Uploaded data file
        self.dict_saveAndSend = {
            'save':{
                'dir':None, 
                'subDir':None            
            }, 
            'send':{
                'to':None, 
                'cc':None, 
                'bcc':None, 
                'subject':None, 
                'body':None
            }
        } # save and send options
        self.mergeBusy = False
        self.addWidgets()

    def addWidgets(self):
        heading_mailMerger = QtWidgets.QLabel('Mail Merger')
        heading_mailMerger.setStyleSheet(HHHconf.design_textMedium)
        heading_mailMerger.setMaximumHeight(20)
        self.vBox_contents.addWidget(heading_mailMerger)

        hBox_top = QtWidgets.QHBoxLayout()
        hBox_top.setContentsMargins(0, 0, 5, 0)
        hBox_top.setSpacing(5)
        hBox_top.setAlignment(QtCore.Qt.AlignTop)  
        self.vBox_contents.addLayout(hBox_top)

        self.mailMergePageList = QtWidgets.QListWidget()
        self.mailMergePageList.setFocusPolicy(QtCore.Qt.NoFocus)
        self.mailMergePageList.setStyleSheet(HHHconf.design_listTwo)
        self.mailMergePageList.setSelectionMode(QtWidgets.QAbstractItemView.NoSelection)
        self.mailMergePageList.setVerticalScrollMode(QtWidgets.QAbstractItemView.ScrollPerPixel)
        self.vBox_contents.addWidget(self.mailMergePageList)

        # Create main mail merge sections
        self.dict_mailMergerStack = {
            'uploadDocuments':['Upload Documents and Templates', self.activateUploadDocumentsSection], 
            'uploadMailMergeData':['Upload Merge Data', self.activateUploadDataSection], 
            'completeMailMerge':['Merge Options', self.activateCompleteMergeSection]
        }
        self.dict_listDataHeaders = {
            'header_variablesFound':['Document Variables', 150], 
            'header_assignColumn':['Assign to Column', 160], 
            'header_defaultValues':['Value if Missing', 190], 
            'header_excludeFromMerge':['Excl.', None]            
        }
        for widget, options in self.dict_mailMergerStack.items():
            item = QtWidgets.QListWidgetItem()
            item.setData(QtCore.Qt.UserRole, widget)
            itemWidget = QtWidgets.QWidget(objectName='itemWidget_' + widget)
            itemLayout = QtWidgets.QVBoxLayout()
            itemLayout.setContentsMargins(0, 0, 0, 0)
            itemLayout.setSpacing(0)
            itemLayout.setAlignment(QtCore.Qt.AlignTop)  
            itemWidget.setLayout(itemLayout)

            currentButton = QtWidgets.QPushButton(options[0], objectName='button_' + widget, clicked=self.selectSection, autoDefault=False)
            if options[1] is not None:
                currentButton.clicked.connect(options[1])
            currentButton.setStyleSheet(HHHconf.design_smallButtonTransparent)
            currentButton.setFixedHeight(50)
            currentButton.setFocusPolicy(QtCore.Qt.NoFocus)
            itemLayout.addWidget(currentButton)

            currentWidget = QtWidgets.QWidget(objectName='childWidget_' + widget)
            currentWidget.setStyleSheet('QWidget#childWidget_' + widget + ' {background-color: ' + HHHconf.secondaryBackgroundColor + ';border: 4px inset #363738}')
            currentWidget.setFixedHeight(600)
            currentWidget.setMaximumHeight(600)
            itemLayout.addWidget(currentWidget)

            item.setSizeHint(itemWidget.sizeHint())
            item.setFlags(item.flags() & ~QtCore.Qt.ItemIsSelectable)
            self.mailMergePageList.addItem(item)
            self.mailMergePageList.setItemWidget(item, itemWidget)

        for widget in self.dict_mailMergerStack:
            self.mailMergePageList.findChild(QtWidgets.QWidget, 'childWidget_' + widget).setHidden(True)
        

        # Section [0]: Upload Documents
        vBox_uploadDocs = QtWidgets.QVBoxLayout()
        vBox_uploadDocs.setContentsMargins(10, 10, 10, 10)
        vBox_uploadDocs.setSpacing(0)
        vBox_uploadDocs.setAlignment(QtCore.Qt.AlignTop)  
        self.mailMergePageList.findChild(QtWidgets.QWidget, 'childWidget_' + list(self.dict_mailMergerStack.keys())[0]).setLayout(vBox_uploadDocs)

        self.uploadDocumentsList = QtWidgets.QListWidget()
        self.uploadDocumentsList.setFocusPolicy(QtCore.Qt.NoFocus)
        self.uploadDocumentsList.setStyleSheet(HHHconf.design_listTwo)
        self.uploadDocumentsList.setSelectionMode(QtWidgets.QAbstractItemView.NoSelection)
        self.uploadDocumentsList.setVerticalScrollMode(QtWidgets.QAbstractItemView.ScrollPerPixel)
        vBox_uploadDocs.addWidget(self.uploadDocumentsList)

        # another document buttons
        hBox_anotherOne = QtWidgets.QHBoxLayout()
        hBox_anotherOne.setContentsMargins(0, 0, 0, 0)
        hBox_anotherOne.setSpacing(5)
        hBox_anotherOne.setAlignment(QtCore.Qt.AlignTop)  
        vBox_uploadDocs.addLayout(hBox_anotherOne)

        addAnotherOneButton = QtWidgets.QPushButton(' +  another one', objectName='button_addAnotherOne', pressed=self.anotherOnePopUp, released=self.anotherOnePopUp, autoDefault=False)
        addAnotherOneButton.clicked.connect(lambda: self.anotherOne(self.DocIndex + 1))
        addAnotherOneButton.setStyleSheet(HHHconf.design_smallButtonTransparent)
        addAnotherOneButton.setFixedHeight(20)
        addAnotherOneButton.setFocusPolicy(QtCore.Qt.NoFocus)
        hBox_anotherOne.addWidget(addAnotherOneButton)

        subtractAnotherOneButton = QtWidgets.QPushButton('- another one', objectName='button_subtractAnotherOne', pressed=self.anotherOnePopUp, released=self.anotherOnePopUp, autoDefault=False)
        subtractAnotherOneButton.clicked.connect(lambda: self.anotherOne(self.DocIndex-1))
        subtractAnotherOneButton.setStyleSheet(HHHconf.design_smallButtonTransparent)
        subtractAnotherOneButton.setFixedHeight(20)
        subtractAnotherOneButton.setFocusPolicy(QtCore.Qt.NoFocus)
        hBox_anotherOne.addWidget(subtractAnotherOneButton)
        subtractAnotherOneButton.setHidden(True)

        # First document upload area - show by default
        self.anotherOne(0)

        # Section [1]: Upload Data
        vBox_uploadData = QtWidgets.QVBoxLayout()
        vBox_uploadData.setContentsMargins(10, 10, 10, 10)
        vBox_uploadData.setSpacing(5)
        vBox_uploadData.setAlignment(QtCore.Qt.AlignTop)  
        self.mailMergePageList.findChild(QtWidgets.QWidget, 'childWidget_' + list(self.dict_mailMergerStack.keys())[1]).setLayout(vBox_uploadData)

        uploadDataArea = HHHconf.dragDropButton(parent=self.mailMergePageList.findChild(QtWidgets.QWidget, 'childWidget_' + list(self.dict_mailMergerStack.keys())[1]), objectName='button_uploadDataArea', text='Browse or Drop Data File Here', dragText='Drop to upload', droppedConnect=self.fileDropped_data)
        uploadDataArea.clicked.connect(self.fileBrowsed_data)
        uploadDataArea.setFixedHeight(100)
        vBox_uploadData.addWidget(uploadDataArea)

        button_clearDataUpload = QtWidgets.QPushButton('Clear Uploaded Data File', objectName='button_clear_data', clicked=self.clearUploadedData, autoDefault=False)
        button_clearDataUpload.setFixedHeight(20)
        button_clearDataUpload.setStyleSheet(HHHconf.design_smallButtonTransparent)
        button_clearDataUpload.setFocusPolicy(QtCore.Qt.NoFocus)
        vBox_uploadData.addWidget(button_clearDataUpload) 

        self.columnVariablesTable = QtWidgets.QTableWidget()
        self.columnVariablesTable.setFocusPolicy(QtCore.Qt.NoFocus)
        self.columnVariablesTable.setStyleSheet(HHHconf.design_table + ' QTableWidget:item{padding: 7 5 5 5;}')
        self.columnVariablesTable.setShowGrid(False)
        self.columnVariablesTable.clearSelection()
        self.columnVariablesTable.setColumnCount(len(self.dict_listDataHeaders.keys()))
        self.columnVariablesTable.setHorizontalHeaderLabels([item[0] for item in self.dict_listDataHeaders.values()])
        # self.columnVariablesTable.horizontalHeader().setMaximumSectionSize(300)
        # self.columnVariablesTable.setMinimumWidth(300)           
        # self.columnVariablesTable.setMaximumWidth(300)
        # self.columnVariablesTable.horizontalHeaderItem(0).setTextAlignment(QtCore.Qt.AlignLeft)
        for i in range(self.columnVariablesTable.columnCount()):
            self.columnVariablesTable.horizontalHeaderItem(i).setTextAlignment(QtCore.Qt.AlignLeft)
            try:
                self.columnVariablesTable.setColumnWidth(i, list(self.dict_listDataHeaders.values())[i][1])
            except:
                self.columnVariablesTable.setColumnWidth(i, 1)
        self.columnVariablesTable.horizontalHeader().setHighlightSections(0)
        self.columnVariablesTable.horizontalHeader().setStyleSheet(HHHconf.design_tableHeader)
        self.columnVariablesTable.horizontalHeader().setStretchLastSection(True)
        self.columnVariablesTable.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Fixed)
        # self.columnVariablesTable.horizontalHeader().setSectionResizeMode(0, QtWidgets.QHeaderView.ResizeToContents)
        # self.columnVariablesTable.horizontalHeader().setSectionResizeMode(2, QtWidgets.QHeaderView.ResizeToContents)
        self.columnVariablesTable.verticalHeader().setVisible(0)
        self.columnVariablesTable.verticalScrollBar().setStyleSheet(HHHconf.design_tableVerticalScrollBar)
        self.columnVariablesTable.setVerticalScrollMode(QtWidgets.QAbstractItemView.ScrollPerPixel)
        self.columnVariablesTable.setSelectionMode(QtWidgets.QAbstractItemView.NoSelection)
        self.columnVariablesTable.setFocusPolicy(QtCore.Qt.NoFocus)
        self.columnVariablesTable.setEditTriggers(QtWidgets.QTableWidget.NoEditTriggers)
        vBox_uploadData.addWidget(self.columnVariablesTable)

        # Section [2]:Complete Merge
        vBox_completeMerge = QtWidgets.QVBoxLayout()
        vBox_completeMerge.setContentsMargins(10, 10, 10, 10)
        vBox_completeMerge.setSpacing(10)
        vBox_completeMerge.setAlignment(QtCore.Qt.AlignTop)  
        self.mailMergePageList.findChild(QtWidgets.QWidget, 'childWidget_' + list(self.dict_mailMergerStack.keys())[2]).setLayout(vBox_completeMerge)

        vBox_completeMergeDocuments = QtWidgets.QVBoxLayout()
        vBox_completeMergeDocuments.setContentsMargins(0, 0, 0, 0)
        vBox_completeMergeDocuments.setSpacing(5)
        vBox_completeMergeDocuments.setAlignment(QtCore.Qt.AlignTop)  
        vBox_completeMerge.addLayout(vBox_completeMergeDocuments)

        self.heading_documents = QtWidgets.QLabel('Documents')
        self.heading_documents.setStyleSheet(HHHconf.design_textMedium)
        self.heading_documents.setMaximumHeight(20)
        vBox_completeMergeDocuments.addWidget(self.heading_documents)

        hBox_saveTo = QtWidgets.QHBoxLayout()
        hBox_saveTo.setContentsMargins(5, 5, 10, 5)
        hBox_saveTo.setSpacing(5)
        hBox_saveTo.setAlignment(QtCore.Qt.AlignTop)  
        vBox_completeMergeDocuments.addLayout(hBox_saveTo)

        self.heading_saveTo = QtWidgets.QLabel('Save To:')
        self.heading_saveTo.setStyleSheet(HHHconf.design_textMedium)
        self.heading_saveTo.setMaximumHeight(20)
        hBox_saveTo.addWidget(self.heading_saveTo)

        self.dict_saveAndSend['save']['dir'] = str(os.path.join(os.environ['USERPROFILE'], 'Desktop\Merge')) # Default dir to save all merged files to
        self.button_saveTo = QtWidgets.QPushButton(self.dict_saveAndSend['save']['dir'], clicked=self.updateSaveDir, default=True)
        self.button_saveTo.setFixedSize(285, 20)
        self.button_saveTo.setStyleSheet(HHHconf.design_smallButtonTransparent)
        hBox_saveTo.addWidget(self.button_saveTo) 

        self.heading_slash = QtWidgets.QLabel('\\')
        self.heading_slash.setStyleSheet(HHHconf.design_textMedium)
        self.heading_slash.setMaximumHeight(20)
        hBox_saveTo.addWidget(self.heading_slash)

        self.combo_saveTo = QtWidgets.QComboBox()
        self.combo_saveTo.setFixedHeight(20)
        self.combo_saveTo.setFixedWidth(155)
        self.combo_saveTo.setStyleSheet(HHHconf.design_comboBox)
        self.combo_saveTo.currentIndexChanged.connect(self.updateSaveSubDir)
        self.combo_saveTo.setEditable(True)
        self.combo_saveTo.lineEdit().setPlaceholderText('Variable folder')
        self.combo_saveTo.lineEdit().setReadOnly(True)
        hBox_saveTo.addWidget(self.combo_saveTo)
        
        hBox_saveTo.addStretch(1)

        self.dict_completeMergeDocsTable = {
            'saveAsName':['\'Save As\' Name', 200], 
            'uniqueSuffix':['Variable Suffix', 150], 
            'saveAsExtension':['File Type', 80], 
            'saveFile':['Save', 40], 
            'attachToEmail':['Attach', 40], 
            'print':['Print', None]           
        }
        self.completeMergeDocsTable = QtWidgets.QTableWidget()
        self.completeMergeDocsTable.setFocusPolicy(QtCore.Qt.NoFocus)
        self.completeMergeDocsTable.setStyleSheet(HHHconf.design_table + ' QTableWidget:item{padding: 7 5 5 5}')
        self.completeMergeDocsTable.setShowGrid(False)
        self.completeMergeDocsTable.clearSelection()
        self.completeMergeDocsTable.setColumnCount(len(self.dict_completeMergeDocsTable.keys()))
        self.completeMergeDocsTable.setHorizontalHeaderLabels([item[0] for item in self.dict_completeMergeDocsTable.values()])
        self.completeMergeDocsTable.setFixedHeight(150)
        
        for i in range(self.completeMergeDocsTable.columnCount()):
            self.completeMergeDocsTable.horizontalHeaderItem(i).setTextAlignment(QtCore.Qt.AlignLeft)
            try:
                self.completeMergeDocsTable.setColumnWidth(i, list(self.dict_completeMergeDocsTable.values())[i][1])
            except:
                self.completeMergeDocsTable.setColumnWidth(i, 1)
        self.completeMergeDocsTable.horizontalHeader().setHighlightSections(0)
        self.completeMergeDocsTable.horizontalHeader().setStyleSheet(HHHconf.design_tableHeader)
        self.completeMergeDocsTable.horizontalHeader().setStretchLastSection(True)
        self.completeMergeDocsTable.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Fixed)
        self.completeMergeDocsTable.verticalHeader().setVisible(0)
        self.completeMergeDocsTable.verticalScrollBar().setStyleSheet(HHHconf.design_tableVerticalScrollBar)
        self.completeMergeDocsTable.setVerticalScrollMode(QtWidgets.QAbstractItemView.ScrollPerPixel)
        self.completeMergeDocsTable.setSelectionMode(QtWidgets.QAbstractItemView.NoSelection)
        self.completeMergeDocsTable.setFocusPolicy(QtCore.Qt.NoFocus)
        self.completeMergeDocsTable.setEditTriggers(QtWidgets.QTableWidget.NoEditTriggers)
        vBox_completeMergeDocuments.addWidget(self.completeMergeDocsTable)

        vBox_completeMergeDocuments.addStretch(1) 


        vBox_completeMergeEmail = QtWidgets.QVBoxLayout()
        vBox_completeMergeEmail.setContentsMargins(0, 0, 0, 0)
        vBox_completeMergeEmail.setSpacing(5)
        vBox_completeMergeEmail.setAlignment(QtCore.Qt.AlignTop)  
        vBox_completeMerge.addLayout(vBox_completeMergeEmail)

        self.heading_email = QtWidgets.QLabel('Email')
        self.heading_email.setStyleSheet(HHHconf.design_textMedium)
        self.heading_email.setMaximumHeight(20)
        vBox_completeMergeEmail.addWidget(self.heading_email)

        vBox_emailSendOptions = QtWidgets.QVBoxLayout()
        vBox_emailSendOptions.setContentsMargins(0, 0, 0, 0)
        vBox_emailSendOptions.setSpacing(0)
        vBox_emailSendOptions.setAlignment(QtCore.Qt.AlignTop)  
        vBox_completeMergeEmail.addLayout(vBox_emailSendOptions)

        hBox_emailRecipients = QtWidgets.QHBoxLayout()
        hBox_emailRecipients.setContentsMargins(5, 5, 5, 5)
        hBox_emailRecipients.setSpacing(5)
        hBox_emailRecipients.setAlignment(QtCore.Qt.AlignTop)  
        vBox_emailSendOptions.addLayout(hBox_emailRecipients)

        self.dict_emailSendOptions = {
            'to':['Send To:', 'required'], 
            'cc':['Cc:', 'optional'], 
            'bcc':['Bcc:', 'optional']
        }
        for emailSendOption in self.dict_emailSendOptions.keys():
            heading = QtWidgets.QLabel(self.dict_emailSendOptions[emailSendOption][0], objectName='label_recipient_' + str(emailSendOption))
            heading.setStyleSheet(HHHconf.design_textMedium)
            heading.setMaximumHeight(20)
            hBox_emailRecipients.addWidget(heading)

            combo = QtWidgets.QComboBox()
            combo.setObjectName('combo_recipient_' + str(emailSendOption))
            combo.setFixedHeight(20)
            combo.setFixedWidth(125)
            combo.currentIndexChanged.connect(self.updateRecipients)
            combo.setStyleSheet(HHHconf.design_comboBox)
            combo.setEditable(True)
            combo.lineEdit().setPlaceholderText(self.dict_emailSendOptions[emailSendOption][1])
            combo.lineEdit().setReadOnly(True)
            hBox_emailRecipients.addWidget(combo)

        hBox_emailRecipients.addStretch(1)

        hBox_emailSubject = QtWidgets.QHBoxLayout()
        hBox_emailSubject.setContentsMargins(5, 5, 5, 5)
        hBox_emailSubject.setSpacing(5)
        hBox_emailSubject.setAlignment(QtCore.Qt.AlignTop)  
        vBox_emailSendOptions.addLayout(hBox_emailSubject)

        self.heading_emailSubject = QtWidgets.QLabel('Subject:')
        self.heading_emailSubject.setStyleSheet(HHHconf.design_textMedium)
        self.heading_emailSubject.setMaximumHeight(20)
        hBox_emailSubject.addWidget(self.heading_emailSubject)

        self.edit_emailSubject =QtWidgets.QLineEdit()
        self.edit_emailSubject.setStyleSheet(HHHconf.design_editBoxTwo)
        self.edit_emailSubject.setFixedWidth(470)
        self.edit_emailSubject.setAlignment(QtCore.Qt.AlignVCenter)
        self.edit_emailSubject.setPlaceholderText('Right-click to insert a column variable')
        self.edit_emailSubject.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.edit_emailSubject.customContextMenuRequested.connect(self.emailSubjectMenu)
        self.edit_emailSubject.textChanged.connect(self.updateSubject)
        hBox_emailSubject.addWidget(self.edit_emailSubject)

        hBox_emailSubject.addStretch(1)

        self.richText_emailBody = HHHconf.HHHTextEditWidget(parent=self, placeholderText='Paste the email\'s body here - formatted text is accepted. Right-click to insert a column variable.')
        self.richText_emailBody.textChanged.connect(self.updateEmailBody)
        self.richText_emailBody.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.richText_emailBody.customContextMenuRequested.connect(self.emailBodyMenu)
        vBox_completeMergeEmail.addWidget(self.richText_emailBody)

        hBox_emailSend = QtWidgets.QHBoxLayout()
        hBox_emailSend.setContentsMargins(0, 0, 0, 0)
        hBox_emailSend.setSpacing(5)
        hBox_emailSend.setAlignment(QtCore.Qt.AlignTop)  
        vBox_completeMergeEmail.addLayout(hBox_emailSend)

        self.userEmailAddress = HHHfunc.mailMergeEngine.PrimarySmtpAddress()

        self.button_testMerge = QtWidgets.QPushButton('Test send to ' + self.userEmailAddress, default=True)
        self.button_testMerge.setFixedHeight(20)
        self.button_testMerge.setFixedWidth(480)
        self.button_testMerge.setStyleSheet(HHHconf.design_smallButtonTransparent)
        self.button_testMerge.clicked.connect(lambda: self.beginMerge(test=True))
        self.button_testMerge.setFocusPolicy(QtCore.Qt.NoFocus)
        hBox_emailSend.addWidget(self.button_testMerge)

        hBox_emailSend.addSpacing(5)

        self.edit_numberOfTestEmails =QtWidgets.QLineEdit()
        self.edit_numberOfTestEmails.setStyleSheet(HHHconf.design_editBoxTwo.replace('padding-left:20px', ''))
        self.edit_numberOfTestEmails.setMaximumWidth(10)
        self.edit_numberOfTestEmails.setPlaceholderText(str(self.defaultNumberOfTestEmails))        
        self.edit_numberOfTestEmails.setAlignment(QtCore.Qt.AlignCenter)
        self.edit_numberOfTestEmails.setValidator(QtGui.QRegExpValidator(QtCore.QRegExp('\d{0,2}')))
        self.edit_numberOfTestEmails.textChanged.connect(self.updateNumberOfTestEmails)
        hBox_emailSend.addWidget(self.edit_numberOfTestEmails)

        self.heading_numberOfTestEmails = QtWidgets.QLabel('emails')
        self.heading_numberOfTestEmails.setStyleSheet(HHHconf.design_textMedium)
        self.heading_numberOfTestEmails.setMaximumHeight(20)
        hBox_emailSend.addWidget(self.heading_numberOfTestEmails)

        self.button_completeMerge = QtWidgets.QPushButton('Complete and Merge', autoDefault=False)
        self.button_completeMerge.setFixedHeight(45)
        self.button_completeMerge.setStyleSheet(HHHconf.design_smallButtonTransparent)
        self.button_completeMerge.clicked.connect(lambda: self.beginMerge())
        self.button_completeMerge.setFocusPolicy(QtCore.Qt.NoFocus)
        vBox_completeMerge.addWidget(self.button_completeMerge)

        vBox_completeMerge.addStretch(1) 
        
        self.mailMergePageList.findChild(QtWidgets.QWidget, 'button_' + list(self.dict_mailMergerStack.keys())[0]).click() # show first section by default

    def selectSection(self):
        buttonClicked = self.sender()
        objName = str(buttonClicked.objectName()).replace('button_', '')
        for widget in self.dict_mailMergerStack:
            if widget != objName:
                self.mailMergePageList.findChild(QtWidgets.QWidget, 'childWidget_' + widget).setHidden(True)
            else:
                self.mailMergePageList.findChild(QtWidgets.QWidget, 'childWidget_' + widget).setHidden(False)
        for i in range(self.mailMergePageList.count()):
            currentItem = self.mailMergePageList.item(i)
            widget = currentItem.data(QtCore.Qt.UserRole)
            currentItem.setSizeHint(self.mailMergePageList.findChild(QtWidgets.QWidget, 'itemWidget_' + widget).sizeHint())   

    def activateUploadDocumentsSection(self):
        pass

    def activateUploadDataSection(self):
        # Re-load table if any new variable discovered
        checkVarList = self.retrieveCurrentDocumentVariables()
        if any(var not in self.dict_foundVars.keys() for var in checkVarList) or any(var not in checkVarList for var in self.dict_foundVars.keys()):
                self.dict_foundVars = dict.fromkeys(checkVarList)
                self.populateDataTable()
                return

    def activateCompleteMergeSection(self):
        # Do not activate unless first two screens deemed valid - if there are docs uploaded with variables, those variables must be assigned a column unless the user has checked them off as excluded.
        self.populateDocumentsTable()

        print('---------------------------------------------------------------------------------')
        print(self.dict_uploadedDocs)
        print('')
        print(self.dict_uploadedData)
        print('')
        print(self.dict_foundVars)
        print('')
        print(self.dict_saveAndSend)
        print('---------------------------------------------------------------------------------')

    def retrieveCurrentDocumentVariables(self):
        foundVars = []
        for uploadedDoc in self.dict_uploadedDocs.values():
            foundVars = foundVars + uploadedDoc['docVars'] # join Lists together
        #Removes Duplicates and creates dictionary keys (values = None)
        foundVarsList = list(set(foundVars))
        foundVarsList.sort(key=lambda v: (v.upper(), v[0].islower()))
        return foundVarsList

    # Section 1 (Upload Docs) methods 
    def fileDropped_template(self, event, sender):
        if event.mimeData().hasUrls():
            for index, url in enumerate(event.mimeData().urls()):
                if index == 0:
                    self.processUploadedDocument(str(url.toLocalFile()), str(sender))
            event.acceptProposedAction()
        else:
            super().dropEvent(event)
    
    def fileBrowsed_template(self):
        fileSelected = QtWidgets.QFileDialog.getOpenFileName(self, 'Select File', 'c:\\', 'Microsoft Word Files (*.doc *.docx *.docm)')
        if len(fileSelected[0]) > 0:
            self.processUploadedDocument(str(fileSelected[0]), str(self.sender().objectName()))

    def processUploadedDocument(self, document, senderButtonName):
        docPath, docExtension = os.path.splitext(str(document))
        docName = os.path.basename(document)
        docNameNoExt, docExt = os.path.splitext(docName)
        objName = senderButtonName.replace('button_uploadTemplateArea_', '')

        try: 
            uploadedDocument = docx.Document(document)
            checkDocForVars = True
        except:
            HHHconf.widgetEffects.flashMessage(self.heading_status, duration=3000, message='File uploaded but not scanned for variables', messageType='warning') 
            checkDocForVars = False

        # Copy file to temp dir and use copy for processing
        try:
            os.makedirs(self.tempMergeDir, exist_ok=True) # Make sure temp directory exists
            documentCopy = shutil.copy(document, self.tempMergeDir + docNameNoExt + '_' + objName + docExt)
        except:
            self.clearUploadedDocument(passed_objName=objName)
            return HHHconf.widgetEffects.flashMessage(self.heading_status, duration=3000, message='File not compatible', messageType='error')

        if len(docName) > 40:
            docName = docName[:20] + '...' + docName[-20:]
        self.uploadDocumentsList.findChild(QtWidgets.QPushButton, senderButtonName).setText(docName)

        # Display widgets
        currentListWidget = self.uploadDocumentsList.findChild(QtWidgets.QListWidget, 'variablesList_' + objName)
        currentListWidget.clear()        
        currentListWidget.setHidden(False)
        self.uploadDocumentsList.findChild(QtWidgets.QPushButton, 'button_clear_' + objName).setHidden(False)
        self.uploadDocumentsList.findChild(QtWidgets.QLabel, 'heading_VariablesFound_' + objName).setHidden(False) 

        varList=[]
        if checkDocForVars:
            # Find variables in doc (headers, footers and main body)
            varList = self.foundVarsInObject(uploadedDocument, runningList=varList)
            for sctn in uploadedDocument.sections:
                varList = self.foundVarsInObject(sctn.header, runningList=varList)
                varList = self.foundVarsInObject(sctn.footer, runningList=varList)
                if sctn.different_first_page_header_footer:
                    varList = self.foundVarsInObject(sctn.first_page_header, runningList=varList)
                    varList = self.foundVarsInObject(sctn.first_page_footer, runningList=varList)
                if uploadedDocument.settings.odd_and_even_pages_header_footer: 
                    # Odd headers/footers captured by .header/.footer objects by default 
                    varList = self.foundVarsInObject(sctn.even_page_header, runningList=varList)
                    varList = self.foundVarsInObject(sctn.even_page_footer, runningList=varList)
            # Display foundVars
            for var in varList:
                currentListWidget.addItem(str(var))
            HHHconf.widgetEffects.flashMessage(self.heading_status, duration=3000, message='Document successfully uploaded', messageType='success') 
            
        # Add/update uploaded document to dictionary for processing later
        self.dict_uploadedDocs[objName] = {
            'path':str(documentCopy), 
            'docVars':varList, 
            'objName':None, 
            'mergeOptions':{
                'saveAsName':docNameNoExt, 
                'uniqueSuffix':None, 
                'saveAsExtension':docExt, 
                'saveFile':False, 
                'attachToEmail':False, 
                'print': False                  
            }
        }
        # Keeps items in same order on the complete mail merge page
        self.dict_uploadedDocs = dict(sorted(self.dict_uploadedDocs.items()))

    def foundVarsInObject(self, obj, runningList=None):
        if runningList is None:
            runningList = []
        newFoundVars = {}
        for pgh in obj.paragraphs:
            newFoundVars = set(re.findall('<<(.+?)>>', pgh.text)) - set(runningList)
            runningList = runningList + list(newFoundVars)
        for tbl in obj.tables:
            for row in tbl.rows:  
                for cell in row.cells: # captures tables in tables 
                        runningList = self.foundVarsInObject(cell, runningList=runningList)
        return list(set(runningList))

    def clearUploadedDocument(self, passed_objName=None):
        if not passed_objName:
            try:
                objName = self.sender().objectName().replace('button_clear_', '')
            except:
                return
        else:
            objName = passed_objName

        # Delete temp file generated on upload
        deletedDoc = self.dict_uploadedDocs.pop(objName, None)
        if deletedDoc is not None:
            deletedDocPath = deletedDoc['path']
            try:
                os.remove(deletedDocPath)
            except Exception as exc:
                pass
        
        self.uploadDocumentsList.findChild(QtWidgets.QPushButton, 'button_uploadTemplateArea_' + objName).setText('Browse or Drop Document File Here')
        self.uploadDocumentsList.findChild(QtWidgets.QListWidget, 'variablesList_' + objName).setHidden(True)
        self.uploadDocumentsList.findChild(QtWidgets.QListWidget, 'variablesList_' + objName).clear()
        self.uploadDocumentsList.findChild(QtWidgets.QPushButton, 'button_clear_' + objName).setHidden(True)
        self.uploadDocumentsList.findChild(QtWidgets.QLabel, 'heading_VariablesFound_' + objName).setHidden(True)

    def anotherOne(self, buttonIndex=0):
        if self.sender().objectName() == 'button_subtractAnotherOne':
            objName = 'uploadDocument_' + str(self.DocIndex)
            self.clearUploadedDocument(passed_objName=objName)
            self.uploadDocumentsList.takeItem(self.DocIndex)            
            self.DocIndex = buttonIndex            
            if buttonIndex == 0:
                self.sender().setHidden(True)
        else:
            if buttonIndex > 10:
                return        
            if buttonIndex == 0:
                self.DocIndex = 0
            else:
                self.findChild(QtWidgets.QPushButton, 'button_subtractAnotherOne').setHidden(False)
                self.DocIndex = buttonIndex
            objName = 'uploadDocument_' + str(buttonIndex)
            item = QtWidgets.QListWidgetItem()
            item.setData(QtCore.Qt.UserRole, objName)
            itemWidget = QtWidgets.QWidget(objectName='widget_' + objName)
            itemWidget.setStyleSheet('background-color:transparent;' + HHHconf.border)
            itemLayout = QtWidgets.QHBoxLayout()
            itemLayout.setContentsMargins(5, 0, 5, 5)
            itemLayout.setSpacing(5)
            itemLayout.setAlignment(QtCore.Qt.AlignTop)  
            itemWidget.setLayout(itemLayout)

            uploadTemplateArea = HHHconf.dragDropButton(parent=self.uploadDocumentsList, objectName='button_uploadTemplateArea_' + objName, text='Browse or Drop Document File Here', dragText='Drop to upload', droppedConnect=self.fileDropped_template)
            uploadTemplateArea.clicked.connect(self.fileBrowsed_template)
            uploadTemplateArea.setFixedHeight(100)
            itemLayout.addWidget(uploadTemplateArea)

            vBox_variablesList = QtWidgets.QVBoxLayout()
            vBox_variablesList.setContentsMargins(0, 0, 0, 0)
            vBox_variablesList.setSpacing(0)
            vBox_variablesList.setAlignment(QtCore.Qt.AlignTop)  
            itemLayout.addLayout(vBox_variablesList)

            heading_variablesFound = QtWidgets.QLabel(self.dict_listDataHeaders['header_variablesFound'][0] + ':', objectName='heading_VariablesFound_' + objName)
            heading_variablesFound.setStyleSheet(HHHconf.design_textSmall)
            heading_variablesFound.setMaximumHeight(18)
            vBox_variablesList.addWidget(heading_variablesFound)

            uploadTemplateVariablesList = QtWidgets.QListWidget()
            uploadTemplateVariablesList.setObjectName('variablesList_' + objName)
            uploadTemplateVariablesList.setFocusPolicy(QtCore.Qt.NoFocus)
            uploadTemplateVariablesList.setStyleSheet(HHHconf.design_listTwo)
            uploadTemplateVariablesList.setSelectionMode(QtWidgets.QAbstractItemView.NoSelection)
            uploadTemplateVariablesList.setVerticalScrollMode(QtWidgets.QAbstractItemView.ScrollPerPixel)
            uploadTemplateVariablesList.setFixedHeight(65)
            uploadTemplateVariablesList.setFixedWidth(200)
            vBox_variablesList.addWidget(uploadTemplateVariablesList)

            button_clearUpload = QtWidgets.QPushButton('Clear', objectName='button_clear_' + objName, clicked=self.clearUploadedDocument, autoDefault=False)
            button_clearUpload.setFixedHeight(20)
            button_clearUpload.setFixedWidth(200)
            button_clearUpload.setStyleSheet(HHHconf.design_smallButtonTransparent)
            button_clearUpload.setFocusPolicy(QtCore.Qt.NoFocus)
            vBox_variablesList.addWidget(button_clearUpload) 

            heading_variablesFound.setHidden(True)
            uploadTemplateVariablesList.setHidden(True)
            button_clearUpload.setHidden(True)

            item.setSizeHint(itemWidget.sizeHint())
            item.setFlags(item.flags() & ~QtCore.Qt.ItemIsSelectable)
            self.uploadDocumentsList.addItem(item)
            self.uploadDocumentsList.setItemWidget(item, itemWidget)

    def anotherOnePopUp(self):
        if self.sender().objectName() == 'button_addAnotherOne':
            if self.findChild(QtWidgets.QPushButton, 'button_addAnotherOne').isDown():
                self.anotherOneWidget = QtWidgets.QWidget(self)
                self.anotherOneLabel = QtWidgets.QLabel(self.anotherOneWidget)
                pix = QtGui.QPixmap(HHHconf.working_dir + '\\icons\\addAnotherOne.jpeg')
                pix = pix.scaledToWidth(100)
                self.anotherOneLabel.setPixmap(pix)
                self.anotherOneWidget.setGeometry(self.sender().geometry().x(), self.sender().geometry().y(), pix.size().width(), pix.size().height())
                self.anotherOneWidget.show()
            else:
                try:
                    self.anotherOneWidget.hide()
                except:
                    pass
        elif self.sender().objectName() == 'button_subtractAnotherOne':
            if self.findChild(QtWidgets.QPushButton, 'button_subtractAnotherOne').isDown():
                self.anotherOneWidget = QtWidgets.QWidget(self)
                self.anotherOneLabel = QtWidgets.QLabel(self.anotherOneWidget)
                pix = QtGui.QPixmap(HHHconf.working_dir + '\\icons\\subtractAnotherOne.jpg')
                pix = pix.scaledToWidth(100)
                self.anotherOneLabel.setPixmap(pix)
                self.anotherOneWidget.setGeometry(self.sender().geometry().x(), self.sender().geometry().y(), pix.size().width(), pix.size().height())
                self.anotherOneWidget.show()
            else:
                try:
                    self.anotherOneWidget.hide()
                except:
                    pass
    
    # Section 2 (Upload Excel Data) methods
    def fileDropped_data(self, event, sender):
        if event.mimeData().hasUrls():
            for index, url in enumerate(event.mimeData().urls()):
                if index == 0:
                    self.processUploadedData(str(url.toLocalFile()), str(sender))
            event.acceptProposedAction()
        else:
            super().dropEvent(event)
    
    def fileBrowsed_data(self):
        fileSelected = QtWidgets.QFileDialog.getOpenFileName(self, 'Select File', 'c:\\', 'Microsoft Excel Files (*.csv *.xls *.xlsx *.xlsm)')
        if len(fileSelected[0]) > 0:
            self.processUploadedData(str(fileSelected[0]), str(self.sender().objectName()))
    
    def processUploadedData(self, document, senderButtonName):
        docName = os.path.basename(document)
        docNameNoExt, docExt = os.path.splitext(docName)

        try: 
            if docExt == '.csv':
                df_uploadedData = pandas.read_csv(document)
            else:
                df_uploadedData = pandas.read_excel(document)
        except:
            self.clearUploadedData()
            return HHHconf.widgetEffects.flashMessage(self.heading_status, duration=3000, message='Error: Only Excel data files are excepted', messageType='error') 

        # convert df columns to correspond with excel columns for easy identification of missing column headers
        df_uploadedData.columns = ['column:' + HHHfunc.numberToXLCol(df_uploadedData.columns.get_loc(col) + 1) if 'Unnamed:' in col else col for col in df_uploadedData.columns]

        if len(docName) > 40:
            docName = docName[:20] + '...' + docName[-20:]
        self.mailMergePageList.findChild(QtWidgets.QPushButton, senderButtonName).setText(docName)

        # Copy file to temp dir and use copy for processing
        try:
            os.makedirs(self.tempMergeDir, exist_ok=True) # Make sure temp directory exists
            documentCopy = shutil.copy(document, self.tempMergeDir + docNameNoExt + '_uploadData' + docExt)
        except:
            self.clearUploadedData()
            return

        self.dict_uploadedData['path'] = documentCopy
        self.dict_uploadedData['dataframe'] = df_uploadedData

        self.updateDataTable()
        self.mailMergePageList.findChild(QtWidgets.QPushButton, 'button_uploadDataArea').setText(docName)
        HHHconf.widgetEffects.flashMessage(self.heading_status, duration=3000, message='Success: Data file uploaded and contents extracted - ' +  str(len(df_uploadedData))  + ' rows detected', messageType='success') 

    def clearUploadedData(self):
        self.mailMergePageList.findChild(QtWidgets.QPushButton, 'button_uploadDataArea').setText('Browse or Drop Data File Here')

        # Delete temp data file
        try:
            os.remove(self.dict_uploadedData['path'])
        except:
            pass
        self.dict_uploadedData = dict.fromkeys(self.dict_uploadedData)   
        for child in self.columnVariablesTable.findChildren(QtWidgets.QLineEdit):
            child.setPlaceholderText('')
            child.setText('')
            child.setAlignment(QtCore.Qt.AlignVCenter)
            child.setReadOnly(False)
        for child in self.columnVariablesTable.findChildren(QtWidgets.QCheckBox):
            child.setCheckState(False)
        for child in self.columnVariablesTable.findChildren(QtWidgets.QComboBox):
            child.setCurrentIndex(-1)  
            child.clear()
            child.lineEdit().setPlaceholderText('No columns found')
                  
        # Reset variable options in var dictionary 
        for currentVarOption in self.dict_foundVars.values():
            currentVarOption['excludeFromMerge'] = False
            currentVarOption['defaultValue'] = ''
            currentVarOption['assignColumn'] = None

    def populateDataTable(self):
        currentRow = 0
        self.columnVariablesTable.setRowCount(currentRow)

        if not self.dict_foundVars: #Display default text if no variables found 
            return

        # #Display found variables in List
        self.columnVariablesTable.setRowCount(len(self.dict_foundVars.keys()))
        for currentVar in self.dict_foundVars.keys():
            # Set up options for each variable
            objName = 'varFound_' + str(currentRow)
            self.dict_foundVars[currentVar] = {'objName':objName, 'assignColumn':None, 'defaultValue':'', 'excludeFromMerge':False}

            label_currentVar = QtWidgets.QLabel(currentVar, objectName='label_' + objName)
            label_currentVar.setStyleSheet(HHHconf.design_textSmall.replace('AlignLeft', '\'AlignVCenter | AlignLeft\''))
            label_currentVar.setFixedHeight(20)
            label_currentVar.setFixedWidth(self.dict_listDataHeaders['header_variablesFound'][1]-10)
            self.columnVariablesTable.setCellWidget(currentRow, 0, label_currentVar)

            combo_assignToColumn = HHHconf.tableComboBox(scrollWidget=self.columnVariablesTable, objectName='combo_' + objName)
            combo_assignToColumn.setFixedHeight(20)
            combo_assignToColumn.setFixedWidth(self.dict_listDataHeaders['header_assignColumn'][1]-10)
            combo_assignToColumn.setStyleSheet(HHHconf.design_comboBox)
            combo_assignToColumn.currentIndexChanged.connect(self.updateAssignColumn)
            combo_assignToColumn.setEditable(True)
            combo_assignToColumn.lineEdit().setPlaceholderText('No columns found')
            combo_assignToColumn.lineEdit().setReadOnly(True)
            self.columnVariablesTable.setCellWidget(currentRow, 1, combo_assignToColumn)

            edit_defaultValue =QtWidgets.QLineEdit(objectName='edit_' + objName)
            edit_defaultValue.setStyleSheet(HHHconf.design_editBoxTwo.replace('padding-left:20px', 'padding-left:10px'))
            edit_defaultValue.setFixedWidth(self.dict_listDataHeaders['header_defaultValues'][1]-10)
            edit_defaultValue.setAlignment(QtCore.Qt.AlignVCenter)
            edit_defaultValue.textChanged.connect(self.updateDefaultValue)
            edit_defaultValue.editingFinished.connect(lambda: self.updateDefaultValue(editingFinished=True))
            self.columnVariablesTable.setCellWidget(currentRow, 2, edit_defaultValue)

            check_excludeVar = QtWidgets.QCheckBox(objectName='check_' + objName)
            check_excludeVar.setStyleSheet(HHHconf.design_checkboxCross + ' QCheckBox{margin-left: 5px}')
            check_excludeVar.stateChanged.connect(self.updateExcludedVars)
            self.columnVariablesTable.setCellWidget(currentRow, 3, check_excludeVar)

            currentRow  += 1

        if self.dict_uploadedData['dataframe'] is not None:
            self.updateDataTable()

    def updateExcludedVars(self):
        objName = self.sender().objectName().replace('check_', '')
        for varOptions in self.dict_foundVars.values():
            if varOptions['objName'] == objName:
                varOptions['excludeFromMerge'] = self.sender().isChecked()
    
    def updateDefaultValue(self, editingFinished=False):
        objName = self.sender().objectName().replace('edit_', '')
        for varOptions in self.dict_foundVars.values():
            if varOptions['objName'] == objName:
                varOptions['defaultValue'] = self.sender().text()
        if editingFinished == True:
            self.sender().setCursorPosition(0)

    def updateAssignColumn(self):
        objName = self.sender().objectName().replace('combo_', '')
        selectedColumn = self.sender().currentText()
        for varOptions in self.dict_foundVars.values():
            if varOptions['objName'] == objName:
                if len(selectedColumn) > 0:
                    varOptions['assignColumn'] = selectedColumn
                else:
                    varOptions['assignColumn'] = None            
        self.sender().lineEdit().setCursorPosition(0)

        # Also update status of the lineEdit
        currentEdit = self.columnVariablesTable.findChild(QtWidgets.QLineEdit, 'edit_' + objName)
        ColumnInDataframeAndDataComplete = self.completeDataFoundinCol(selectedColumn)
        if ColumnInDataframeAndDataComplete is True:
            currentEdit.setPlaceholderText('data complete')
            currentEdit.setText('')
            currentEdit.setAlignment(QtCore.Qt.AlignBottom)
            currentEdit.setReadOnly(True)
        elif ColumnInDataframeAndDataComplete is False:
            currentEdit.setPlaceholderText('')
            currentEdit.setText('')
            currentEdit.setReadOnly(False)
        else:
            currentEdit.setPlaceholderText('')
            currentEdit.setText('')
            currentEdit.setReadOnly(True)

    def completeDataFoundinCol(self, columnName):
        if self.dict_uploadedData['dataframe'] is not None:
            df_uploadedData = self.dict_uploadedData['dataframe']
            if columnName not in list(df_uploadedData):
                return None
            if columnName in df_uploadedData.columns[df_uploadedData.isna().any()].tolist():
                return False
        return True

    def updateDataTable(self):
        if self.dict_uploadedData['dataframe'] is None:
            print('no data found')
            return
        
        df_uploadedData = self.dict_uploadedData['dataframe']

        for row in range(self.columnVariablesTable.rowCount()):
            currentVar = self.columnVariablesTable.cellWidget(row, 0).text()
            currentCombo = self.columnVariablesTable.cellWidget(row, 1)
            currentEdit = self.columnVariablesTable.cellWidget(row, 2)

            currentCombo.lineEdit().setPlaceholderText('')
            colList = df_uploadedData.columns.values.tolist()
            colList.sort(key=lambda v: (v.upper(), v[0].islower()))
            for col in colList:
                currentCombo.addItem(col)                
            currentCombo.setCurrentIndex(currentCombo.findText(currentVar)) # -1 if not found, which will display placeholder text
            ColumnInDataframeAndDataComplete = self.completeDataFoundinCol(currentVar) 
            if ColumnInDataframeAndDataComplete is True:
                currentEdit.setPlaceholderText('data complete')
                currentEdit.setText('')
                currentEdit.setAlignment(QtCore.Qt.AlignBottom)
                currentEdit.setReadOnly(True)
            elif ColumnInDataframeAndDataComplete is None:
                currentEdit.setPlaceholderText('')
                currentEdit.setText('')
                currentEdit.setAlignment(QtCore.Qt.AlignBottom)
                currentEdit.setReadOnly(False)

    # Section 3 (Complete Merge) methods
    def updateSaveDir(self):
        dirSelected = QtWidgets.QFileDialog.getExistingDirectory(self, 'Select \'Save to\' Directory')
        if len(dirSelected) > 0:
            dirSelected = dirSelected.replace('/', '\\')
            self.dict_saveAndSend['save']['dir'] = dirSelected
            if len(dirSelected) > 36:
                dirSelected = dirSelected[:18] + '...' + dirSelected[-18:]
            self.button_saveTo.setText(dirSelected)

    def updateSaveSubDir(self):
        if len(self.sender().currentText()) > 0:
            self.dict_saveAndSend['save']['subDir'] = self.sender().currentText()
        else:
            self.dict_saveAndSend['save']['subDir'] = None
        self.sender().lineEdit().setCursorPosition(0)

    def populateDocumentsTable(self):
        if self.dict_uploadedData['dataframe'] is None:
            HHHconf.widgetEffects.flashMessage(self.heading_status, duration=3000, message='Warning: No merge data detected.', messageType='warning') 

        # load columns into list for dropdowns (also used later for dropdowns in table)
        self.colList = []
        if self.dict_uploadedData['dataframe'] is not None:
            df_uploadedData = self.dict_uploadedData['dataframe']
            self.colList = df_uploadedData.columns.values.tolist()
            self.colList.sort(key=lambda v: (v.upper(), v[0].islower()))

        self.combo_saveTo.clear()    
        self.combo_saveTo.addItem('')
        for col in self.colList:
            self.combo_saveTo.addItem(col)                
        self.combo_saveTo.setCurrentIndex(0)

        for recipientCombo in self.dict_emailSendOptions.keys():
            try:
                currentCombo = self.findChild(QtWidgets.QComboBox, 'combo_recipient_' + recipientCombo)
                currentCombo.clear()
                currentCombo.addItem('')
                for col in self.colList:
                    currentCombo.addItem(col)                
                currentCombo.setCurrentIndex(0)
            except:
                pass

        currentRow = 0
        self.completeMergeDocsTable.setRowCount(currentRow)
        if not self.dict_uploadedDocs: #Display default text if no docs found 
            return
        # Display found docs in List
        self.completeMergeDocsTable.setRowCount(len(self.dict_uploadedDocs.keys()))
        for uploadedDoc in self.dict_uploadedDocs.values():
            objName = 'docFound_' + str(currentRow)
            # Provides a pointer to the document in self.dict_uploadedDocs as the user changes the doc options
            uploadedDoc['objName'] = objName

            edit_saveAsName =QtWidgets.QLineEdit(uploadedDoc['mergeOptions']['saveAsName'], objectName='edit_' + objName)
            edit_saveAsName.setCursorPosition(0) # Display long text from left
            edit_saveAsName.setStyleSheet(HHHconf.design_editBoxTwo.replace('padding-left:20px', 'padding-left:10px'))
            edit_saveAsName.setFixedWidth(self.dict_completeMergeDocsTable['saveAsName'][1]-10)
            edit_saveAsName.setAlignment(QtCore.Qt.AlignVCenter)
            edit_saveAsName.textChanged.connect(self.updateSaveAsName)
            self.completeMergeDocsTable.setCellWidget(currentRow, 0, edit_saveAsName)

            combo_uniqueSuffix = HHHconf.tableComboBox(scrollWidget=self.completeMergeDocsTable, objectName='combo_' + objName)
            combo_uniqueSuffix.setFixedHeight(20)
            combo_uniqueSuffix.setFixedWidth(self.dict_completeMergeDocsTable['uniqueSuffix'][1]-10)
            combo_uniqueSuffix.setStyleSheet(HHHconf.design_comboBox)
            combo_uniqueSuffix.currentIndexChanged.connect(self.updateUniqueSuffix)
            combo_uniqueSuffix.setEditable(True)
            combo_uniqueSuffix.lineEdit().setPlaceholderText('optional')
            combo_uniqueSuffix.lineEdit().setReadOnly(True)
            self.completeMergeDocsTable.setCellWidget(currentRow, 1, combo_uniqueSuffix)
            # Load data columns into dropdowns
            combo_uniqueSuffix.addItem('')
            for col in self.colList:
                combo_uniqueSuffix.addItem(col)                
            combo_uniqueSuffix.setCurrentIndex(0)

            # If document is Microsoft Word file, give option to save as PDF 
            if uploadedDoc['mergeOptions']['saveAsExtension'] in ['.doc', '.docx', '.docm', '.pdf']:
                combo_saveAsExtension = HHHconf.tableComboBox(scrollWidget=self.completeMergeDocsTable, objectName='labelOrCombo_' + objName)
                combo_saveAsExtension.setFixedHeight(20)
                combo_saveAsExtension.setFixedWidth(self.dict_completeMergeDocsTable['saveAsExtension'][1]-10)
                combo_saveAsExtension.setStyleSheet(HHHconf.design_comboBox)
                combo_saveAsExtension.currentIndexChanged.connect(self.updateSaveAsExtension)
                combo_saveAsExtension.setEditable(True)
                combo_saveAsExtension.lineEdit().setPlaceholderText('None')
                combo_saveAsExtension.lineEdit().setReadOnly(True)
                combo_saveAsExtension.addItem(uploadedDoc['mergeOptions']['saveAsExtension'])
                if uploadedDoc['mergeOptions']['saveAsExtension'] != '.pdf':
                    combo_saveAsExtension.addItem('.pdf')
                combo_saveAsExtension.setCurrentIndex(0)
                self.completeMergeDocsTable.setCellWidget(currentRow, 2, combo_saveAsExtension)
            else:
                label_saveAsExtension = QtWidgets.QLabel(uploadedDoc['mergeOptions']['saveAsExtension'], objectName='labelOrCombo_' + objName)
                label_saveAsExtension.setStyleSheet(HHHconf.design_textSmall.replace('AlignLeft', '\'AlignVCenter | AlignLeft\''))
                label_saveAsExtension.setFixedHeight(20)
                label_saveAsExtension.setFixedWidth(self.dict_completeMergeDocsTable['saveAsExtension'][1]-10)
                self.completeMergeDocsTable.setCellWidget(currentRow, 2, label_saveAsExtension)

            check_saveFile = QtWidgets.QCheckBox(objectName='checkSave_' + objName)
            check_saveFile.setStyleSheet(HHHconf.design_checkboxTick + ' QCheckBox{margin-left: 5px}')
            check_saveFile.stateChanged.connect(self.updateSaveFile)
            self.completeMergeDocsTable.setCellWidget(currentRow, 3, check_saveFile)

            check_emailAttach = QtWidgets.QCheckBox(objectName='checkAttach_' + objName)
            check_emailAttach.setStyleSheet(HHHconf.design_checkboxTick + ' QCheckBox{margin-left: 5px}')
            check_emailAttach.stateChanged.connect(self.updateAttachEmail)
            self.completeMergeDocsTable.setCellWidget(currentRow, 4, check_emailAttach)

            check_print = QtWidgets.QCheckBox(objectName='checkPrint_' + objName)
            check_print.setStyleSheet(HHHconf.design_checkboxTick + ' QCheckBox{margin-left: 5px}')
            check_print.stateChanged.connect(self.updatePrint)
            self.completeMergeDocsTable.setCellWidget(currentRow, 5, check_print)

            currentRow  += 1

    def updateSaveAsName(self):
        objName = self.sender().objectName().replace('edit_', '')
        for docOptions in self.dict_uploadedDocs.values():
            if docOptions['objName'] == objName:
                docOptions['mergeOptions']['saveAsName'] = self.sender().text()

    def updateUniqueSuffix(self):
        objName = self.sender().objectName().replace('combo_', '')
        for docOptions in self.dict_uploadedDocs.values():
            if docOptions['objName'] == objName:
                if len(self.sender().currentText()) > 0:
                    docOptions['mergeOptions']['uniqueSuffix'] = self.sender().currentText()
                else:
                    docOptions['mergeOptions']['uniqueSuffix'] = None
        self.sender().lineEdit().setCursorPosition(0)

    def updateSaveAsExtension(self):
        objName = self.sender().objectName().replace('labelOrCombo_', '')
        for docOptions in self.dict_uploadedDocs.values():
            if docOptions['objName'] == objName:
                if len(self.sender().currentText()) > 0:
                    docOptions['mergeOptions']['saveAsExtension'] = self.sender().currentText()
                else:
                    docOptions['mergeOptions']['saveAsExtension'] = None

    def updateSaveFile(self):
        objName = self.sender().objectName().replace('checkSave_', '')
        for docOptions in self.dict_uploadedDocs.values():
            if docOptions['objName'] == objName:
                 docOptions['mergeOptions']['saveFile'] = self.sender().isChecked()

    def updateAttachEmail(self):
        objName = self.sender().objectName().replace('checkAttach_', '')
        for docOptions in self.dict_uploadedDocs.values():
            if docOptions['objName'] == objName:
                 docOptions['mergeOptions']['attachToEmail'] = self.sender().isChecked()

    def updatePrint(self):
        objName = self.sender().objectName().replace('checkPrint_', '')
        for docOptions in self.dict_uploadedDocs.values():
            if docOptions['objName'] == objName:
                 docOptions['mergeOptions']['print'] = self.sender().isChecked()

    def updateRecipients(self):
        objName = self.sender().objectName().replace('combo_recipient_', '')
        txt = self.sender().currentText()
        if len(txt) > 0:
            # Validate if email column is data complete
            if not self.completeDataFoundinCol(txt):
                HHHconf.widgetEffects.flashMessage(self.heading_status, duration=3000, message='Warning: Missing recipients detected - those emails will not be sent', messageType='warning') 
            self.dict_saveAndSend['send'][objName] = txt
        else:
            self.dict_saveAndSend['send'][objName] = None
        self.sender().lineEdit().setCursorPosition(0)

    def updateSubject(self):
        if len(self.sender().text()) > 0:
            self.dict_saveAndSend['send']['subject'] = self.sender().text()
        else:
            self.dict_saveAndSend['send']['subject'] = None

    def updateEmailBody(self):
        if len(self.sender().toPlainText()) > 0:
            self.sender().setStyleSheet(HHHconf.design_textEdit.replace(HHHconf.monospaceFont, '').replace(HHHconf.mainFontColor, 'black').replace(HHHconf.mainBackgroundColor, '#fafafa') + HHHconf.design_tableVerticalScrollBar)
            self.dict_saveAndSend['send']['body'] = self.sender().toHtml()
        else:
            self.sender().setStyleSheet(HHHconf.design_textEdit + HHHconf.design_tableVerticalScrollBar)
            self.dict_saveAndSend['send']['body'] = None            

    def emailSubjectMenu(self):
        self.insertColList = []
        for column in self.colList:
            current_col = {}
            current_col['type'] = 'action'
            current_col['displayText'] = column
            current_col['checkable'] = False
            current_col['checkedState'] = False
            current_col['connect'] = self.insertColIntoEmailSubject
            self.insertColList.append(current_col)

        menuItems = [
                {'type':'menu', 'displayText': 'Insert Column Variable', 'subMenu': self.insertColList}
            ]

        emailSubjectMenu = HHHconf.CustomMenu(menuItems)
        # emailSubjectMenu.signal_checkableAction.connect() # Used for checkable menu items
        emailSubjectMenu.menuAction = emailSubjectMenu.exec_(QtGui.QCursor.pos())

        for item in menuItems:
            if item['type'] == 'action':
                if emailSubjectMenu.menuAction == item['object'] and item['connect'] is not None:
                    item['connect']()
                    return
        for item in self.insertColList:
            if item['type'] == 'action':
                if emailSubjectMenu.menuAction == item['object'] and item['connect'] is not None:
                    item['connect'](item['displayText'])
                    return

    def insertColIntoEmailSubject(self, displayText=None):
        self.edit_emailSubject.insert('<<' + displayText + '>>')

    def emailBodyMenu(self):
        self.insertColList = []
        for column in self.colList:
            current_col = {}
            current_col['type'] = 'action'
            current_col['displayText'] = column
            current_col['checkable'] = False
            current_col['checkedState'] = False
            current_col['connect'] = self.insertColIntoEmailBody
            self.insertColList.append(current_col)

        menuItems = [
                {'type':'menu', 'displayText': 'Insert Column Variable', 'subMenu': self.insertColList}, 
                {'type':'section', 'displayText': None, 'checkable': False, 'checkedState': False, 'connect': None}, 
                {'type':'action', 'displayText': 'Undo', 'checkable': False, 'checkedState': False, 'connect':self.richText_emailBody.undo}, 
                {'type':'action', 'displayText': 'Redo', 'checkable': False, 'checkedState': False, 'connect':self.richText_emailBody.redo}, 
                {'type':'section', 'displayText': None, 'checkable': False, 'checkedState': False, 'connect': None}, 
                {'type':'action', 'displayText': 'Copy', 'checkable': False, 'checkedState': False, 'connect':self.richText_emailBody.copy}, 
                {'type':'action', 'displayText': 'Cut', 'checkable': False, 'checkedState': False, 'connect':self.richText_emailBody.cut}, 
                {'type':'action', 'displayText': 'Paste', 'checkable': False, 'checkedState': False, 'connect':self.richText_emailBody.paste}
            ]

        emailBodyMenu = HHHconf.CustomMenu(menuItems)
        # emailBodyMenu.signal_checkableAction.connect() # Used for checkable menu items
        emailBodyMenu.menuAction = emailBodyMenu.exec_(QtGui.QCursor.pos())

        for item in menuItems:
            if item['type'] == 'action':
                if emailBodyMenu.menuAction == item['object'] and item['connect'] is not None:
                    item['connect']()
                    return
        for item in self.insertColList:
            if item['type'] == 'action':
                if emailBodyMenu.menuAction == item['object'] and item['connect'] is not None:
                    item['connect'](item['displayText'])
                    return
    
    def insertColIntoEmailBody(self, displayText=None):
        self.richText_emailBody.insertPlainText('<<' + displayText + '>>')        

    def updateNumberOfTestEmails(self):
        # Resize to contents
        charCount = len(self.sender().text())
        self.sender().setFixedWidth(min(10*(charCount + 1), 20)) 
        if charCount > 0:
            self.numberOfTestEmails = int(self.sender().text())
        else:
            self.numberOfTestEmails = self.defaultNumberOfTestEmails
        
    def beginMerge(self, test=False):
        # return if a merge is already occurring
        if self.mergeBusy:
            return HHHconf.widgetEffects.flashMessage(label=self.heading_status, duration=3000, message='A merge is already being executed in the background. Please wait for completion or restart the program', messageType='warning')
        # Validate - All merges requre data      
        if self.dict_uploadedData['dataframe'] is None:
            return
        # Validate - if no docs uploaded, and no email elements detected, merge should not occur
        if not self.dict_uploadedDocs and all(x is None for x in self.dict_saveAndSend['send'].values()):
            return HHHconf.widgetEffects.flashMessage(self.heading_status, duration=3000, message='Mimimum merge requirements not fulfilled', messageType='error')
        # Validate - if docs uploaded but not elected to be saved, attached to email, or printed, then docs are not ready
        if self.dict_uploadedDocs:
            for uploadedDocOptions in self.dict_uploadedDocs.values():
                if not any(x is True for x in uploadedDocOptions['mergeOptions'].values()):
                    return HHHconf.widgetEffects.flashMessage(self.heading_status, duration=3000, message='Not all uploaded documents have been selected for processing', messageType='error')
        # if any doc variables exist, validate whether they have been assigned data columns to merge with
        if self.dict_foundVars:
            for docVarOptions in self.dict_foundVars.values():
                if docVarOptions['excludeFromMerge'] is False:
                    if docVarOptions['assignColumn'] is None:
                        return HHHconf.widgetEffects.flashMessage(self.heading_status, duration=3000, message='Not all included document variables have been assigned a data column', messageType='error')
                    elif docVarOptions['defaultValue'] == '':
                        HHHconf.widgetEffects.flashMessage(self.heading_status, duration=3000, message='Warning: Not all included document variables have been given default values', messageType='warning')

        # Create local copies of data to alter for execution, and preserve actual user-inputs
        dict_uploadedData = copy.deepcopy(self.dict_uploadedData)
        dict_saveAndSend = copy.deepcopy(self.dict_saveAndSend)
        dict_uploadedDocs = copy.deepcopy(self.dict_uploadedDocs)
        dict_foundVars = copy.deepcopy(self.dict_foundVars)

        # TEST MERGE
        if test is True:
            # Validate if email is ready to send - has mail merge data, recipient column, email subject, email body
            if any(x is None for x in [dict_uploadedData['dataframe'], dict_saveAndSend['send']['to'], dict_saveAndSend['send']['subject'], dict_saveAndSend['send']['body']]):
                return HHHconf.widgetEffects.flashMessage(self.heading_status, duration=3000, message='Minimum email requirements not fulfilled', messageType='error')

            # add a single line to top of email body stating recipients if it weren't a test - uses the master dict
            emailTestString = 'This is a test email. Intended recipients for the actual merge are as follows: To: &lt;&lt;_emailTo_&gt;&gt; Cc: &lt;&lt;_emailCc_&gt;&gt; Bcc: &lt;&lt;_emailBcc_&gt;&gt; <br>'

            emailBody = dict_saveAndSend['send']['body']
            pos = emailBody.find('>', emailBody.find('<body'))
            dict_saveAndSend['send']['body'] = emailBody[:pos + 1] + emailTestString + emailBody[pos + 1:]

            # limit number of tests to user selection
            df_mergeDataToExecute = dict_uploadedData['dataframe'].head(self.numberOfTestEmails)
            # create columns if df for retaining original recipient info during the test merge
            df_mergeDataToExecute['_emailTo_'] = df_mergeDataToExecute[dict_saveAndSend['send']['to']]
            try:
                df_mergeDataToExecute['_emailCc_'] = df_mergeDataToExecute[dict_saveAndSend['send']['cc']]
            except:
                pass
            try:
                df_mergeDataToExecute['_emailBcc_'] = df_mergeDataToExecute[dict_saveAndSend['send']['bcc']]
            except:
                pass
            # Change the chosen recipient column to user's email address
            df_mergeDataToExecute[dict_saveAndSend['send']['to']] = self.userEmailAddress
            dict_uploadedData['dataframe'] = df_mergeDataToExecute
            # remove any cc's and bcc's (AFTER they've been retained in another column) uses the copied dict, to preserve master dict data
            dict_saveAndSend['send']['cc'] = None
            dict_saveAndSend['send']['bcc'] = None
            print(df_mergeDataToExecute)
            HHHconf.widgetEffects.flashMessage(self.heading_status, duration=3000, message='Mail merge TEST is now being executed - ' + str(self.numberOfTestEmails) + ' emails will be sent', messageType='success')

        # ACTUAL MERGE
        else:
            # Validate if any email elements have been entered. If so, check it is ready to send - has mail merge data, recipient column, email subject, email body
            if any(x is not None for x in dict_saveAndSend['send'].values()) and any(x is None for x in [dict_uploadedData['dataframe'], dict_saveAndSend['send']['to'], dict_saveAndSend['send']['subject'], dict_saveAndSend['send']['body']]):
                return HHHconf.widgetEffects.flashMessage(self.heading_status, duration=3000, message='Email elements detected, but minimum email requirements not fulfilled', messageType='error')

            # validate password to continue
            if not HHHfunc.mainEngine.requestPassword(parent=self, user=self.username, title='Confirm', dialogText='You are about to begin the mail merge. Please enter your password to continue:'):
                return
            HHHconf.widgetEffects.flashMessage(self.heading_status, duration=3000, message='Mail Merge is now being executed', messageType='success')

        self.mergeBusy = True
        self.signal_mailMergeStarted.emit()
        # begin merge in thread
        threadComplete = threading.Event()
        mergeThread = threading.Thread(target=HHHfunc.mailMergeEngine.start, kwargs={'threadCompleteEvent':threadComplete, 'mergeData':dict_uploadedData, 'mergeOptions':dict_saveAndSend, 'mergeDocuments':dict_uploadedDocs, 'varToCols':dict_foundVars, 'username':self.username})
        completionWaitThread = threading.Thread(target=self.completeMergeEvent, kwargs={'threadCompleteEvent':threadComplete}, daemon=True)
        mergeThread.start()
        completionWaitThread.start()
    
    def completeMergeEvent(self, threadCompleteEvent):
        if threadCompleteEvent.wait():
            self.mergeBusy = False
            self.signal_mailMergeComplete.emit()