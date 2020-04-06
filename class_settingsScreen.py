import pandas, datetime, re, openpyxl, docx, shutil, threading, copy

from argon2 import PasswordHasher
from PyQt5 import QtCore, QtGui, QtWidgets

import HHHconf, HHHfunc

class settingsScreen(HHHconf.HHHWindowWidget):

    signal_settingsUpdated = QtCore.pyqtSignal(object)

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.addWidgets()
        self.get_userCredentialsTable()
        
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
        loginDetailsBox = QtWidgets.QGroupBox()
        loginDetailsBox.setContentsMargins(5, 5, 5, 5)
        loginDetailsBox.setStyleSheet(HHHconf.design_GroupBox)
        self.vBox_contents.addWidget(loginDetailsBox)

        vBox_loginDetailsBox = QtWidgets.QVBoxLayout()
        vBox_loginDetailsBox.setContentsMargins(5, 5, 5, 5)
        vBox_loginDetailsBox.setSpacing(5)
        vBox_loginDetailsBox.setAlignment(QtCore.Qt.AlignTop)   
        loginDetailsBox.setLayout(vBox_loginDetailsBox) 

        heading_loginDetails = QtWidgets.QLabel('Login Details')
        heading_loginDetails.setStyleSheet(HHHconf.design_textMedium)
        heading_loginDetails.setMaximumHeight(20)
        vBox_loginDetailsBox.addWidget(heading_loginDetails)

        loginDetailsSpacer = QtWidgets.QVBoxLayout()
        loginDetailsSpacer.setContentsMargins(50, 5, 50, 5)
        loginDetailsSpacer.setSpacing(5)
        loginDetailsSpacer.setAlignment(QtCore.Qt.AlignTop)   
        vBox_loginDetailsBox.addLayout(loginDetailsSpacer) 

        dict_loginDetailsWidgets = { # objectName:[headingText, readOnly, textChanged, validator, button, defaultValue]
            HHHconf.userAdmin_username:['Username', True, None, None, None, self.username], 
            HHHconf.userAdmin_latestLogin:['Logged in since', True, None, None, None, None], 
            HHHconf.userAdmin_latestPasswordUpdate:['Password updated', True, None, None, ['updateLoginPassword', 'Update your password', self.updateLoginPassword], None], 
            HHHconf.userAdmin_firstName:['First Name', True, None, None, None, None], 
            HHHconf.userAdmin_lastName:['Last Name', True, None, None, None, None], 
            HHHconf.userAdmin_emailAddress:['Email Address', False, None, None, None, None], 
            HHHconf.userAdmin_authority:['User Authority', True, None, None, None, None]
        }

        for widget, options in dict_loginDetailsWidgets.items():
            hBox_currentLoginDetail = QtWidgets.QHBoxLayout()
            hBox_currentLoginDetail.setContentsMargins(0, 0, 0, 0)
            hBox_currentLoginDetail.setSpacing(5)
            loginDetailsSpacer.addLayout(hBox_currentLoginDetail)
            
            currentHeading = QtWidgets.QLabel(options[0], objectName='heading_loginDetails_' + widget)
            currentHeading.setFixedWidth(115)
            currentHeading.setStyleSheet(HHHconf.design_textSmall)
            currentHeading.setFixedHeight(20)
            hBox_currentLoginDetail.addWidget(currentHeading)

            currentWidget = QtWidgets.QLineEdit('', objectName='object_loginDetails_' + widget)
            currentWidget.setStyleSheet(HHHconf.design_editBoxTwo)
            currentWidget.setAlignment(QtCore.Qt.AlignBottom)
            currentWidget.setFixedHeight(20)
            currentWidget.setPlaceholderText('No Data Found')
            currentWidget.setReadOnly(options[1])
            if options[2] is not None:
                currentWidget.textChanged.connect(options[2])
            if options[3] is not None:
                currentWidget.setValidator(options[3])
            hBox_currentLoginDetail.addWidget(currentWidget)
            if options[4] is not None:
                currentWidget.setFixedWidth(200)
                currentWidgetButton = QtWidgets.QPushButton(options[4][1], objectName='button_' + options[4][0], clicked=options[4][2], default=True)
                currentWidgetButton.setStyleSheet(HHHconf.design_smallButtonTransparent)
                currentWidgetButton.setFixedHeight(20)
                hBox_currentLoginDetail.addWidget(currentWidgetButton)
                hBox_currentLoginDetail.addStretch(1)
            if options[5] is not None:
                currentWidget.setText(str(options[5]))

        self.getLoginDetails(self.username)

        # User Credentials hBox
        hBox_userCredentials = QtWidgets.QHBoxLayout()
        hBox_userCredentials.setContentsMargins(0, 0, 0, 0)
        hBox_userCredentials.setSpacing(5)
        hBox_userCredentials.setAlignment(QtCore.Qt.AlignTop)   
        self.vBox_contents.addLayout(hBox_userCredentials) 

        vBox_userCredentialsTable = QtWidgets.QVBoxLayout()
        vBox_userCredentialsTable.setContentsMargins(0, 0, 0, 0)
        vBox_userCredentialsTable.setSpacing(5)
        vBox_userCredentialsTable.setAlignment(QtCore.Qt.AlignTop)  
        hBox_userCredentials.addLayout(vBox_userCredentialsTable)

        # User Credentials Table 
        self.heading_appCredentials = QtWidgets.QLabel('App Credentials')
        self.heading_appCredentials.setStyleSheet(HHHconf.design_textMedium)
        vBox_userCredentialsTable.addWidget(self.heading_appCredentials)
        self.userCredentialsTable = HHHconf.HHHTableWidget(contextMenu=self.userCredentialsTableMenu, doubleClicked=self.updateUserCredentials)
        vBox_userCredentialsTable.addWidget(self.userCredentialsTable)
        self.button_update.clicked.connect(self.updateSettings)

        # User Authority Matrix (access for HHH_maximum user level only)
        if 'changeUserAuthorities' in HHHconf.dict_permissions[self.df_loginDetails.at[0, HHHconf.userAdmin_authority]]:
            vBox_userAuthoritiesMatrix = QtWidgets.QVBoxLayout()
            vBox_userAuthoritiesMatrix.setContentsMargins(0, 0, 0, 0)
            vBox_userAuthoritiesMatrix.setSpacing(5)
            vBox_userAuthoritiesMatrix.setAlignment(QtCore.Qt.AlignTop)  
            self.vBox_contents.addLayout(vBox_userAuthoritiesMatrix)

            self.heading_userAuthorityMatrix = QtWidgets.QLabel('Change User Authorities')
            self.heading_userAuthorityMatrix.setStyleSheet(HHHconf.design_textMedium)
            vBox_userAuthoritiesMatrix.addWidget(self.heading_userAuthorityMatrix)

            # TODO
            # User Dropdown List
            # Authority Dropdown List
            # Add New Authority Type Screen


    def getLoginDetails(self, username):
        self.df_loginDetails = HHHfunc.mainEngine.getLoginDetails(username)
        self.currentPasswordHashed = self.df_loginDetails.at[0, HHHconf.userAdmin_password]

        for child in self.findChildren(QtWidgets.QLineEdit):
            try:
                data = self.df_loginDetails.at[0, str(child.objectName()).replace('object_loginDetails_', '')]
                if pandas.core.dtypes.common.is_datetime_or_timedelta_dtype(self.df_loginDetails[str(child.objectName()).replace('object_loginDetails_', '')]):
                    child.setText(datetime.datetime.strftime(data, HHHconf.dateFormat2 + ' %I:%M:%S%p'))
                # elif str(child.objectName()).replace('object_loginDetails_', '') == HHHconf.userAdmin_authority:
                #     child.setText(data.replace('HHH_', ''))   
                else:
                    child.setText(data)         
            except: 
                pass

    def updateLoginPassword(self):
        self.childWindow = passwordUpdateScreen(currentPassword=self.currentPasswordHashed, username=self.username, geometry=self.geometry(), padding={'left':50, 'top':50, 'right':50, 'bottom':self.geometry().height() / 4}, title='Update Password', updateButton='Update password', parent=self)
        self.childWindow.show()
        self.childWindow.activateWindow()
        self.childWindow.setFocus()
        self.childWindow.signal_passwordUpdated.connect(lambda: self.getLoginDetails(self.username))
        self.childWindow.signal_passwordUpdated.connect(lambda x: self.signal_settingsUpdated.emit({HHHconf.events_EventTime: datetime.datetime.now(),HHHconf.events_UserName: self.username, HHHconf.events_Event: HHHconf.dict_eventCategories['settings'], HHHconf.events_EventDescription: str(x)}))
        self.childWindow.signal_passwordUpdated.connect(lambda x: HHHconf.widgetEffects.flashMessage(self.heading_status, duration=3000, message=str(x), messageType='success'))

    def get_userCredentialsTable(self):        
        self.df_userCredentials = HHHfunc.mainEngine.getAppCredentials(self.username)
        self.df_userCredentials.drop(labels=HHHconf.userCredentials_user, axis='columns', inplace=True)
        
        # Decrypt passwords for display (length only - characters hidden)
        self.df_userCredentials[HHHconf.userCredentials_password] = self.df_userCredentials[HHHconf.userCredentials_password].apply(lambda x: HHHfunc.mainEngine.twoWayDecrypt(str(x)))

        self.df_userCredentialsCopy = copy.deepcopy(self.df_userCredentials) # unhidden password required for the editCredentialsScreen

        for col in self.df_userCredentials:
            if col == HHHconf.userCredentials_password:
                self.df_userCredentials[col] = self.df_userCredentials[col].map(lambda x: u'\u25CF' * len(x))
            else:
                self.df_userCredentials[col] = self.df_userCredentials[col].map(lambda x: str(x) if x is not None else '')

        self.df_userCredentials.rename({HHHconf.userCredentials_application:'Application', HHHconf.userCredentials_username:'Username', HHHconf.userCredentials_password:'Password', HHHconf.userCredentials_credential1:'Token 1', HHHconf.userCredentials_credential2:'Token 2'}, axis='columns', inplace=True)
        self.userCredentialsTable.setModel(self.df_userCredentials)
        self.userCredentialsTable.sortByColumn(0, QtCore.Qt.AscendingOrder)

    def userCredentialsTableMenu(self, pos): 
        menu = QtWidgets.QMenu()
        menu.setStyleSheet(HHHconf.design_tableMenu)
        
        action_updateCredentials = menu.addAction('Update credentials')
        action_addCredentials = menu.addAction('Add a new login')
        action_deleteCredentials = menu.addAction('Delete credentials')

        action = menu.exec_(QtGui.QCursor.pos())
        if action == action_updateCredentials and len(self.df_userCredentials.index) > 0:
            self.updateUserCredentials()
        elif action == action_addCredentials:
            self.addUserCredentials()
        elif action == action_deleteCredentials and len(self.df_userCredentials.index) > 0:
            self.deleteUserCredentials()
    
    def updateUserCredentials(self):
        self.childWindow = appCredentialsScreen(app=self.userCredentialsTable.model().index(self.userCredentialsTable.currentIndex().row(), 0).data(QtCore.Qt.DisplayRole), appUser=self.userCredentialsTable.model().index(self.userCredentialsTable.currentIndex().row(), 1).data(QtCore.Qt.DisplayRole), currentPassword=self.df_userCredentialsCopy.loc[self.df_userCredentialsCopy[HHHconf.userCredentials_application] == self.userCredentialsTable.model().index(self.userCredentialsTable.currentIndex().row(), 0).data(QtCore.Qt.DisplayRole), HHHconf.userCredentials_password].iloc[0], username=self.username, geometry=self.geometry(), padding={'left':50, 'top':50, 'right':50, 'bottom':self.geometry().height() / 4}, title='Update App Credentials', updateButton='Update Credentials', parent=self)
        self.childWindow.show()
        self.childWindow.activateWindow()
        self.childWindow.setFocus()
        self.childWindow.signal_passwordUpdated.connect(self.get_userCredentialsTable)
        self.childWindow.signal_passwordUpdated.connect(lambda x: self.signal_settingsUpdated.emit({HHHconf.events_EventTime: datetime.datetime.now(),HHHconf.events_UserName: self.username, HHHconf.events_Event: HHHconf.dict_eventCategories['settings'], HHHconf.events_EventDescription: str(x)}))
        self.childWindow.signal_passwordUpdated.connect(lambda x: HHHconf.widgetEffects.flashMessage(self.heading_status, duration=3000, message=str(x), messageType='success'))

    def addUserCredentials(self):
        self.childWindow = appCredentialsScreen(app=None, appUser=None, currentPassword=None, username=self.username, geometry=self.geometry(), padding={'left':50, 'top':50, 'right':50, 'bottom':self.geometry().height() / 4}, title='Add App Credentials', updateButton='Add Credentials', parent=self)
        self.childWindow.show()
        self.childWindow.activateWindow()
        self.childWindow.setFocus()
        self.childWindow.signal_passwordUpdated.connect(self.get_userCredentialsTable)
        self.childWindow.signal_passwordUpdated.connect(lambda x: self.signal_settingsUpdated.emit({HHHconf.events_EventTime: datetime.datetime.now(),HHHconf.events_UserName: self.username, HHHconf.events_Event: HHHconf.dict_eventCategories['settings'], HHHconf.events_EventDescription: str(x)}))
        self.childWindow.signal_passwordUpdated.connect(lambda x: HHHconf.widgetEffects.flashMessage(self.heading_status, duration=3000, message=str(x), messageType='success'))

    def deleteUserCredentials(self):
        appName = str(self.userCredentialsTable.model().index(self.userCredentialsTable.currentIndex().row(), 0).data(QtCore.Qt.DisplayRole))
        if not HHHfunc.mainEngine.requestPassword(self, self.username, title='Confirm Action', dialogText='You are about to delete the credentials to ' + appName + '. Please enter your password to continue:'):
            return
        HHHfunc.mainEngine.deleteAppCredentials(username=self.username, app=appName)
        self.get_userCredentialsTable()
        self.signal_settingsUpdated.emit({HHHconf.events_EventTime: datetime.datetime.now(),HHHconf.events_UserName: self.username, HHHconf.events_Event: HHHconf.dict_eventCategories['settings'], HHHconf.events_EventDescription: 'App credentials deleted: ' + str(appName)})
        return HHHconf.widgetEffects.flashMessage(self.heading_status, duration=1000, message='App credentials deleted: ' + str(appName), messageType='success')

    def updateSettings(self):
        # Currently only updates user's email address
        self.findChild(QtWidgets.QLineEdit, 'object_loginDetails_' + HHHconf.userAdmin_emailAddress).clearFocus()
        HHHfunc.mainEngine.updateLoginDetails(self.username, data={HHHconf.userAdmin_emailAddress:self.findChild(QtWidgets.QLineEdit, 'object_loginDetails_' + HHHconf.userAdmin_emailAddress).text()})
        self.getLoginDetails(self.username)
        return HHHconf.widgetEffects.flashMessage(self.heading_status, duration=1000, message='Email address updated', messageType='success')

    def settingsCancel(self):
        self.close()

class passwordUpdateScreen(HHHconf.HHHWindowWidget):

    # Used to update password for MAIN APP (HHH) only
    signal_passwordUpdated = QtCore.pyqtSignal(str)

    def __init__(self, currentPassword, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.currentPassword = None if currentPassword is None else str(currentPassword) 
        self.addWidgets()
        self.pwdhsr = PasswordHasher()
        self.passwordTimer = QtCore.QTimer(timeout=self.validateCurrentPassword)
        self.passwordTimer.setSingleShot(True)        

    def addWidgets(self):
        self.vBox_contents.setSpacing(15)

        dict_passwordChange = { # objectName:[headingText, textChanged, placeholderText]
            'currentPassword':['Current Password', self.validateCurrentPasswordTimer, 'Please enter your current password'], 
            'newPassword':['New Password', self.validateNewPassword, 'Please enter a new password'], 
            'confirmNewPassword':['Confirm Password', self.validateConfirmNewPassword, 'Please re-type the new password']
        }
        for widget, options in dict_passwordChange.items():
            hBox_currentRow = QtWidgets.QHBoxLayout()
            hBox_currentRow.setContentsMargins(0, 0, 0, 0)
            hBox_currentRow.setSpacing(5)
            hBox_currentRow.setAlignment(QtCore.Qt.AlignTop)  
            self.vBox_contents.addLayout(hBox_currentRow)

            currentHeading = QtWidgets.QLabel(options[0], objectName='heading_' + widget)
            currentHeading.setFixedWidth(110)
            currentHeading.setStyleSheet(HHHconf.design_textSmallError)
            currentHeading.setFixedHeight(20)
            hBox_currentRow.addWidget(currentHeading)

            currentWidget = QtWidgets.QLineEdit('', objectName='object_' + widget)
            currentWidget.setStyleSheet(HHHconf.design_editBoxTwo)
            currentWidget.setAlignment(QtCore.Qt.AlignRight)
            currentWidget.setFixedHeight(20)
            currentWidget.textChanged.connect(options[1])
            currentWidget.setPlaceholderText(options[2])
            currentWidget.setEchoMode(QtWidgets.QLineEdit.Password)
            currentWidget.setProperty('isPassword', True)
            hBox_currentRow.addWidget(currentWidget)

        # Final Buttons
        button_showPassword = QtWidgets.QPushButton('Unhide Passwords', objectName='button_showPassword', pressed=self.showPassword, released=self.showPassword, autoDefault=False)
        button_showPassword.setStyleSheet(HHHconf.design_smallButtonTransparent)
        button_showPassword.setFixedHeight(20)
        button_showPassword.setFocusPolicy(QtCore.Qt.NoFocus)
        self.vBox_contents.addWidget(button_showPassword)

        self.button_update.clicked.connect(self.updatePassword)

    def showPassword(self):
        if self.findChild(QtWidgets.QPushButton, 'button_showPassword').isDown():
            for passwordFields in self.findChildren(QtWidgets.QLineEdit):
                passwordFields.setEchoMode(QtWidgets.QLineEdit.Normal)
        else:
            for passwordFields in self.findChildren(QtWidgets.QLineEdit):
                if passwordFields.property('isPassword'):
                    passwordFields.setEchoMode(QtWidgets.QLineEdit.Password)

    def updatePassword(self):
        for validate_status in [self.validateCurrentPassword(), self.validateNewPassword(clearConfirmPassword=False), self.validateConfirmNewPassword()]:
            if validate_status != True:
                return HHHconf.widgetEffects.flashMessage(self.heading_status, duration=1000, message='Validation Error - ' + str(validate_status), messageType='error')

        rawPassword = str(self.findChild(QtWidgets.QLineEdit, 'object_confirmNewPassword').text())
        hashedPassword = HHHfunc.mainEngine.oneWayEncrypt(rawPassword)

        HHHfunc.mainEngine.updateLoginDetails(self.username, data={HHHconf.userAdmin_password:hashedPassword, HHHconf.userAdmin_latestPasswordUpdate:datetime.datetime.now()})
        self.close()
        self.signal_passwordUpdated.emit('HHH password updated')

    def updatePasswordCancel(self):
        self.close()

    def validateCurrentPasswordTimer(self):
        self.passwordTimer.start(300)

    def validateCurrentPassword(self):
        try:
            rawPassword = self.findChild(QtWidgets.QLineEdit, 'object_currentPassword').text()
            self.pwdhsr.verify(self.currentPassword, rawPassword + str(len(rawPassword)))
            self.findChild(QtWidgets.QLabel, 'heading_currentPassword').setStyleSheet(HHHconf.design_textSmallSuccess)
            return True
        except:
            self.findChild(QtWidgets.QLabel, 'heading_currentPassword').setStyleSheet(HHHconf.design_textSmallError)
            return 'currentPassword'
            
    def validateNewPassword(self, enteredPassword=None, clearConfirmPassword=True):
        if clearConfirmPassword is True: # make user reconfirm password if they change 
            self.findChild(QtWidgets.QLineEdit, 'object_confirmNewPassword').setText('')
        if len(self.findChild(QtWidgets.QLineEdit, 'object_newPassword').text()) > 0:
            self.findChild(QtWidgets.QLabel, 'heading_newPassword').setStyleSheet(HHHconf.design_textSmallSuccess)
            self.currentNewPassword = self.findChild(QtWidgets.QLineEdit, 'object_newPassword').text()
            return True
        else:
            self.findChild(QtWidgets.QLabel, 'heading_newPassword').setStyleSheet(HHHconf.design_textSmallError)
            self.currentNewPassword = None
            return 'newPassword'

    def validateConfirmNewPassword(self):
        try:
            if len(self.findChild(QtWidgets.QLineEdit, 'object_confirmNewPassword').text()) > 0 and self.findChild(QtWidgets.QLineEdit, 'object_confirmNewPassword').text() == self.currentNewPassword:
                self.findChild(QtWidgets.QLabel, 'heading_confirmNewPassword').setStyleSheet(HHHconf.design_textSmallSuccess)
                return True
        except:
            pass
        self.findChild(QtWidgets.QLabel, 'heading_confirmNewPassword').setStyleSheet(HHHconf.design_textSmallError)
        return 'confirmNewPassword' 

# Used for adding credentials for an app used in HHH, or updating existing credentials
class appCredentialsScreen(passwordUpdateScreen):
    def __init__(self, app, appUser, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.appUser = appUser
        self.app = app
        # get list of apps not yet registered
        df_appCredentials = HHHfunc.mainEngine.getAppCredentials(self.username)
        list_unregisteredApps = list(set(HHHconf.list_appNames) - set(df_appCredentials[HHHconf.userCredentials_application]))
        # new widgets for this screen
        self.dict_credentialsWidgets = { # objectName:[objectType, headingText, textChanged / indexChanged, placeholderText, defaultValue / comboBoxItems]
            'appName':[QtWidgets.QComboBox(), 'Application', self.validateApp, 'Please select an application', list_unregisteredApps], 
            'appUser':[QtWidgets.QLineEdit(), 'App Username', self.validateappUser, 'Please enter the application username', self.appUser]
        }
        indx = 0
        for widget, options in self.dict_credentialsWidgets.items():
            hBox_currentRow = QtWidgets.QHBoxLayout()
            hBox_currentRow.setContentsMargins(0, 0, 0, 0)
            hBox_currentRow.setSpacing(5)
            hBox_currentRow.setAlignment(QtCore.Qt.AlignTop)  
            self.vBox_contents.insertLayout(indx, hBox_currentRow)

            currentHeading = QtWidgets.QLabel(options[1], objectName='heading_' + widget)
            currentHeading.setFixedWidth(110)
            currentHeading.setStyleSheet(HHHconf.design_textSmallError)
            currentHeading.setFixedHeight(20)
            hBox_currentRow.addWidget(currentHeading)

            if isinstance(options[0], QtWidgets.QLineEdit):
                currentWidget = QtWidgets.QLineEdit('', objectName='object_' + widget)
                currentWidget.setStyleSheet(HHHconf.design_editBoxTwo)
                currentWidget.setAlignment(QtCore.Qt.AlignRight)
                currentWidget.setFixedHeight(20)
                hBox_currentRow.addWidget(currentWidget)
                if options[2] is not None:
                    currentWidget.textChanged.connect(options[2])
                if options[3] is not None:
                    currentWidget.setPlaceholderText(options[3])
                if options[4] is not None:
                    currentWidget.setText(options[4])
                currentWidget.setProperty('isPassword', False)
            elif isinstance(options[0], QtWidgets.QComboBox):
                currentWidget = QtWidgets.QComboBox()
                currentWidget.setObjectName('object_' + widget)
                currentWidget.setStyleSheet(HHHconf.design_comboBox.replace('3px', '20px'))
                hBox_currentRow.addWidget(currentWidget)
                if options[2] is not None:
                        currentWidget.currentIndexChanged.connect(options[2])
                if options[3] is not None:
                    currentWidget.setEditable(True)  
                    currentWidget.lineEdit().setPlaceholderText(options[3])
                if options[4] is not None:
                    for app in options[4]:
                        currentWidget.addItem(str(app))          
                currentWidget.lineEdit().setReadOnly(True)   
                currentWidget.setCurrentIndex(-1)
                if self.app is not None: # Disallow changes to appName if user is updating credentials - only allow if user is adding new app credentials 
                    currentWidget.addItem(self.app)
                    currentWidget.setCurrentIndex(currentWidget.findText(self.app))
                    currentWidget.setEnabled(False)
            indx  += 1
        # Override currentPassword behaviour
        # Display currentPassword in corresponding lineEdit and also disable editing (user is either adding or updating their app credentials, so they should see their current password) 
        self.findChild(QtWidgets.QLineEdit, 'object_currentPassword').setReadOnly(True)
        self.findChild(QtWidgets.QLineEdit, 'object_currentPassword').setPlaceholderText('')
        if self.currentPassword is not None:
            self.findChild(QtWidgets.QLineEdit, 'object_currentPassword').setText(self.currentPassword)
            self.findChild(QtWidgets.QLabel, 'heading_currentPassword').setStyleSheet(HHHconf.design_textSmall)
            self.button_update.setText('Update Credentials')
        else:
            self.findChild(QtWidgets.QLabel, 'heading_currentPassword').setStyleSheet(HHHconf.design_textSmall.replace(HHHconf.mainFontColor, 'gray;text-decoration: line-through'))
            self.findChild(QtWidgets.QLineEdit, 'object_currentPassword').setFocusPolicy(QtCore.Qt.NoFocus)
            self.button_update.setText('Add Credentials')

    # Override old methods in passwordUpdateScreen
    def showPassword(self):
        if self.findChild(QtWidgets.QComboBox, 'object_appName').currentText() == 'HALO' and HHHfunc.mainEngine.verifyPermissions('viewHaloPassword') is False:
            return HHHconf.widgetEffects.flashMessage(self.heading_status, duration=1000, message='Password hidden for security - contact the admin for information', messageType='warning')
        super(appCredentialsScreen, self).showPassword()

    def updatePassword(self):
        for validate_status in [self.validateApp(), self.validateappUser(), self.validateNewPassword(clearConfirmPassword=False), self.validateConfirmNewPassword()]:
            if validate_status != True:
                return HHHconf.widgetEffects.flashMessage(self.heading_status, duration=1000, message='Validation Error - ' + str(validate_status), messageType='error')
        app = self.findChild(QtWidgets.QComboBox, 'object_appName').currentText()
        if app in HHHconf.list_commonAppNames and HHHfunc.mainEngine.verifyPermissions('changeHaloCredentials') is False:
            return HHHconf.widgetEffects.flashMessage(self.heading_status, duration=1000, message='Insufficient Permissions', messageType='warning')
        HHHfunc.mainEngine.updateAppCredentials(app=app, data={HHHconf.userCredentials_username: self.findChild(QtWidgets.QLineEdit, 'object_appUser').text(), HHHconf.userCredentials_password: HHHfunc.mainEngine.twoWayEncrypt(self.findChild(QtWidgets.QLineEdit, 'object_confirmNewPassword').text())}, username=self.username)
        self.close()
        self.signal_passwordUpdated.emit('App credentials updated: ' + str(app))

    def validateCurrentPassword(self): # override original validate method (app credentials do not need current pw validated)
        return True

    # New methods 
    def validateApp(self):
        if self.findChild(QtWidgets.QComboBox, 'object_appName').currentIndex() >= 0:
            self.findChild(QtWidgets.QLabel, 'heading_appName').setStyleSheet(HHHconf.design_textSmallSuccess)
            return True
        self.findChild(QtWidgets.QLabel, 'heading_appName').setStyleSheet(HHHconf.design_textSmallError)
        return 'appName'

    def validateappUser(self):
        if len(self.findChild(QtWidgets.QLineEdit, 'object_appUser').text()) > 0:
            self.findChild(QtWidgets.QLabel, 'heading_appUser').setStyleSheet(HHHconf.design_textSmallSuccess)
            return True
        self.findChild(QtWidgets.QLabel, 'heading_appUser').setStyleSheet(HHHconf.design_textSmallError)
        return 'appUser'