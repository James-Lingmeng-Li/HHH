"""Configuration

This is were certain variables are stored to be used by the app when required, and should be called direct to some screen in the GUI later so that users can edit certain variables. 

This also contains important dictionaries as well as the custom GUI classes used throughout the other modules
"""

import os

from PyQt5 import QtCore, QtGui, QtWidgets

border = 'border: 1px solid transparent;' # Change to red during testing

# Variables
username_PC = os.getenv('username')
stateList = ['NSW', 'QLD', 'SA', 'TAS', 'VIC', 'WA', 'ACT', 'NT']
periodsPerYearList = ['4', '12', '26', '52']
maxOriginalBalance = 1000000.00
maxYears = 5
maxInterestRate = 20.00
subordinateDefaultAmount = 50.00
gstComponent = 0.10
subordinateNotes_systemGenerated = 'System-generated'
widgetPlaceHolderText_notes = 'No notes entered'
status_pendingAdd = 'AUTO_GENERATED'
list_appNames = ['AQUARIUS', 'HALO', 'HALO_DEMO', 'WPC COL', 'ANZ Global']
list_commonAppNames = ['HALO', 'HALO_DEMO']
importReceiptsAQCode = '0004' # Found in AQ SHLNE screen
reverseReceiptsAQCode = '0005' # Found in AQ SHLNE screen
translucifyOpacityLevel = 0.4
defaultNumberofTestEmails = 5 # Mail Merger

sqlServer = 'MTIC2AQDB16'
sqlDatabase = 'mPower'
sqlUsername = 'mpowerusr'
sqlPassword = 'PwusrMT19'

defaultPassword = 'password' # The default password for a new user
haloPassword = 'C0rtana' # The leading substring of the password for Halo 
haloDemoPassword = 'M0n3ytech' # The leading substring of the password for Halo Demo 
working_dir = os.getcwd()
local_dir = 'C:\\Users\\' + username_PC + '\\AppData\\Local\\HHH'
localExceptionLogger_dir = local_dir + '\\exceptionsLog'
icons_dir = working_dir + '\\icons'
scripts_dir = working_dir + '\\library'
downloads_dir = scripts_dir + '\\downloads'
chromeDriver_dir = scripts_dir + '\\webdriver\\chromedriver.exe'
common_dir = '\\\\moneytech.com.au\\common'
fernetKey_dir = common_dir + '\\Customer Filing\\mPower\\initkey.txt'
aq_dir = 'C:\\aq-client\\bin\\aquarius64.exe' # Dev AQ = 'C:\\devaq-client\\bin\\aquarius64.exe'
temp_dir = local_dir + '\\_temp\\_mailMergeDocs\\'
saveSegments_dir = common_dir + '\\Customer Filing\\Debtor Finance\\a Segment Accounts (All)'

ahk_dir = working_dir + '\\AutoHotkey\\AutoHotkeyU64.exe' # OLD - not required
importDebtorReceipt_dir = scripts_dir + '\\importDebtorReceipt.exe' # compiled AHK script. To edit, go to scripts_dir\ahk and edit. Once done, use AHKCompiler to compile into exe
toggleInput_dir = scripts_dir + '\\toggleInput.exe' # compiled AHK script. To edit, go to scripts_dir\ahk and edit .ahk file. Once done, use AHKCompiler to compile into exe

dict_WPCClients = {'Progility':{'segmentLink':'CommsAus', 'contactName':'Accounts', 'recipient':{'to':'arcommsaust@commsaust.com.au; ar@progilitytechnologies.com; Dayan.DeSilva@progilitytechnologies.com', 'cc':None, 'bcc':None}}, 'Fertoz':{'segmentLink':'FERTAG', 'contactName':'Accounts', 'recipient':{'to':'les.szonyi@aus2.com; ttaylor@firstclassaccounts.com', 'cc':None, 'bcc':None}}} # temporary dictionary for our WPC clients - should convert to a config screen on the gui, so account managers can add or remove clients and email recipients
WPCClients = 'Progility:CommsAus, Fertoz:Fertag' # used to convert to AHK (key, value) array
link_wpc = 'https://online.corp.westpac.com.au/'
email_accountsTeam = 'accounts@moneytech.com.au' # set to your own email for dev
printerID = 'FX DocuCentre-V C5576 PCL 6 on MTECHDC08R201.moneytech.com.au'
link_halo = 'https://mtx2.moneytech.com.au/halo-aqlive/' # 'https://mtx2.moneytech.com.au/halo-aqlive/' # DEV HALO
link_haloDemo = 'https://testaq.moneytech.com.au/halo-aqlive/'

dateFormat = 'dd/MM/yyyy'
dateFormat2 = '%d/%m/%Y'
dateFormat3 = '%Y/%m/%d' # for SQL queries
dateFormat4 = '%d-%m-%Y'
dateFormat5 = '%m%y'
moneyFormat = '${:0,.2f}'
timeFormat = '%I:%M:%S %p'
dateTimeFormat = '%Y/%m/%d %I:%M:%S %p'

# Visuals
name_app = 'HHH' # DEV = 'HHH dev'
name_tabHome = 'Home'
name_tabDebtor = 'Debtor'
name_tabEquipment = 'Equipment'   

appLeft = 20
appTop = 40
appWidth = 600 # Currently overriden by background size
appHeight = 600 # Currently overriden by background size

#BUG: Fonts not working correctly
headerFont = 'Roboto Condensed'
labelFont = 'Roboto'
monospaceFont = 'Consolas' # monospace
tableFont = 'Calibri'
smallFontSize = '12px'
mediumFontSize = '14px'
largeFontSize = '15px'
hugeFontSize = '18px'
superFontSize = '24px'
mainBackgroundColor = 'rgba(88, 89, 91, 255)' #58595b with transparency value
secondaryBackgroundColor = '#424345' # darker than mainBackgroundColor 
mainFontColor = '#ffffff'
errorFontColor = '#ff9e9e'
successFontColor = '#a4ff9e'
warningFontColor = '#f7941d' 
editFontColor = '#d1d3d4'
selectedColor = '#f7941d'
hoverColor = '#1e6bb3'
buttonBackgroundColor = '#3c3c3c'
tableHeaderColor = '#3c3c3c'
transparentTableColor = 'rgba(255, 255, 255, 10)'
overlayColor = '#3c3c3c'
transparentWidgetColor = 'rgba(255, 255, 255, 100)'
windowObject = 'windowWidget'

design_background = 'background-color:' + mainBackgroundColor
design_textSmall = 'QLabel{' + border + 'color:' + mainFontColor + ';font:' + smallFontSize + ' ' + labelFont + ';font-weight: 100;qproperty-alignment: AlignLeft}'
design_textSmallWarning = design_textSmall + 'QLabel{color:' + warningFontColor + '}'
design_textSmallError = design_textSmall + 'QLabel{color:' + errorFontColor + '}'
design_textSmallSuccess = design_textSmall + 'QLabel{color:' + successFontColor + '}'
design_textMedium = 'QLabel{' + border + 'color:' + mainFontColor + ';background-color: transparent;font: ' + mediumFontSize + ' ' + labelFont + ';font-weight: 100;qproperty-alignment: AlignLeft}'
design_textMediumError = design_textMedium + 'QLabel{color:' + errorFontColor + '}'
design_textMediumSuccess = design_textMedium + 'QLabel{color:' + successFontColor + '}'
design_textLarge = 'QLabel{' + border + 'color:' + mainFontColor + ';background-color: transparent;font:' + largeFontSize + ' ' + labelFont + ';font-weight: 100;qproperty-alignment: AlignLeft}'
design_textLargeError = design_textLarge + 'QLabel{color:' + errorFontColor + '}'
design_textLargeSuccess = design_textLarge + 'QLabel{color:' + successFontColor + '}'
design_textSmallMonoSpace = 'QLabel{' + border + 'color:' + mainFontColor + ';background-color: transparent;font:' + smallFontSize + ' ' + monospaceFont + ';qproperty-alignment: AlignLeft}'
design_textMediumMonospace = 'QLabel{' + border + 'color:' + mainFontColor + ';background-color: transparent;font:' + mediumFontSize + ' ' + monospaceFont + ';qproperty-alignment: AlignLeft}'
design_textHuge = 'QLabel{' + border + 'color:' + mainFontColor + ';background-color: transparent;font:' + hugeFontSize + ' ' + labelFont + ';qproperty-alignment: AlignHCenter|AlignVCenter}'

design_tabs = 'QTabWidget:pane{border-bottom:1px solid #ffffff; border-top: 1px solid #ffffff;position: absolute} QTabWidget:tab-bar{alignment: center;} QTabBar{font:' + largeFontSize + ' Candara;color:' + mainFontColor + ';outline: 0} QTabBar:tab {background: ' + buttonBackgroundColor + ';font-size:' + largeFontSize + ';width:100px;height:30px;border-right: 10px solid;border-right-color:' + buttonBackgroundColor + ';border-left: 10px solid;border-left-color:' + buttonBackgroundColor + ';border-bottom: 5px solid;border-bottom-color: ' + buttonBackgroundColor + ';} QTabBar::tab:hover {color:' + mainFontColor + ';background: ' + buttonBackgroundColor + ';border-bottom: 5px solid;border-bottom-color: ' + hoverColor + ';} QTabBar::Tab:!selected{margin-top:5px} QTabBar::tab:selected {background: ' + buttonBackgroundColor + ';color:' + selectedColor + ';border-bottom: 5px solid;border-bottom-color:' + selectedColor + ';height:35px}'

design_table = 'QTableView{outline: 0px; background-color: ' + transparentTableColor + '; border: none;color:' + mainFontColor + ';font:' + smallFontSize + ' ' + tableFont + ';selection-background-color:' + selectedColor + ';selection-color:' + mainFontColor + ';gridline-color:gray;} QTableView::item:focus{border: 0px;background-color:' + selectedColor + '} QLineEdit{background-color: ' + mainBackgroundColor + ';font:bold;color:' + editFontColor + '}' 
design_tableHeader = 'QHeaderView::section{background-color:' + tableHeaderColor + ';color:' + mainFontColor + ';font: 11px ' + headerFont + ';qproperty-alignment: \'AlignLeft | AlignHCenter\'; min-height: 25px}'
design_tableVerticalScrollBar = 'QScrollBar:vertical{width:12px;background-color: transparent; border:none; margin: 0px 5px 0px 5px} QScrollBar:hover:vertical{margin: 0px 3px 0px 3px} QScrollBar::handle:vertical{border-radius: 3px; background-color: ' + tableHeaderColor + '; min-height: 50px} QScrollBar::handle:hover:vertical{background-color:' + selectedColor + '} QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical{background:transparent;} QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical{background: none;border: none; height: 0px}'
design_tableHorizontalScrollBar = 'QScrollBar:horizontal{height:12px;background-color: transparent; border:none; margin: 5px 0px 5px 0px} QScrollBar:hover:horizontal{margin: 3px 0px 3px 0px} QScrollBar::handle:horizontal{border-radius: 3px; background-color: ' + tableHeaderColor + '; min-width: 50px} QScrollBar::handle:hover:horizontal{background-color:' + selectedColor + '} QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal{background:transparent;} QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal{background: none;border: none; height: 0px}'
design_tableMenu = 'QMenu{border:2px solid ' + hoverColor + ';background-color:' + mainBackgroundColor + ';color:' + mainFontColor + ';font:' + largeFontSize + ' ' + headerFont + '} QMenu:item{background-color:transparent;border: 1px solid transparent; padding: 2px 20px} QMenu:item:selected{background-color:' + buttonBackgroundColor + ';border-color: ' + hoverColor + '} QMenu:item:checked{color:' + selectedColor + '}'

design_list = 'QListView{color:' + mainFontColor + ';background-color: transparent;border: 0px solid gray} QListView::item:hover{background:' + hoverColor + '} QListView::item:selected{color:' + mainFontColor + ';background:' + selectedColor + '}' + design_tableVerticalScrollBar + design_tableHorizontalScrollBar
design_listReadOnly = 'QListView{color:' + mainFontColor + ';background-color: transparent;border: 0px} QListView::item:hover{background:' + hoverColor + '} QListView::item:selected{color:' + mainFontColor + ';background:' + selectedColor + '}' + design_tableVerticalScrollBar + design_tableHorizontalScrollBar
design_listTwo = 'QListView{color:' + mainFontColor + ';background-color: transparent;' + border + '} QListView::item:hover{background:transparent;' + border + '} QListView::item:selected{color:' + mainFontColor + ';background:transparent}' + design_tableVerticalScrollBar + design_tableHorizontalScrollBar

design_textButton = 'QPushButton{' + border + 'color:' + mainFontColor + ';font:' + largeFontSize + ' ' + headerFont + ';background-color: transparent;padding:2px 2px} QPushButton:hover{color:' + hoverColor + ';text-decoration: underline} QPushButton:pressed{color:' + selectedColor + '}'

design_smallButtonTransparent = 'QPushButton{color:' + mainFontColor + ';font: ' + largeFontSize + ' ' + headerFont + ';background-color: transparent;border-radius:0px;border:1px solid gray;padding-left: 10px; padding-right: 10px} QPushButton:focus{outline:0;border:1px solid;border-color:' + mainFontColor + ';} QPushButton:hover{border:2px solid;border-color:' + hoverColor + ';background-color:' + buttonBackgroundColor + '} QPushButton:pressed{background-color:' + hoverColor + ';border-style: inset;} QPushButton:checked{border: 2px solid ' + selectedColor + '} QPushButton:checked:pressed{background-color:' + selectedColor + ';border-style: inset;}'
design_smallButtonTransparentDrag = 'QPushButton{color:' + mainFontColor + ';font: ' + largeFontSize + ' ' + headerFont + ';background-color: transparent;border-radius:0px;border:1px dashed white;padding-left: 10px; padding-right: 10px} QPushButton:focus{outline:0;border:0.5px dashed;border-color:' + hoverColor + ';} QPushButton:hover{border:2px solid;border-color:' + hoverColor + ';background-color:' + buttonBackgroundColor + '} QPushButton:pressed{background-color:' + hoverColor + ';border-style: inset;}'
design_smallButtonTransparentDragEnter = 'QPushButton{color:' + mainFontColor + ';font: ' + largeFontSize + ' ' + headerFont + ';background-color: transparent;border-radius:0px;border:1px dashed white;padding-left: 10px; padding-right: 10px} QPushButton:focus{outline:0;border:0.5px dashed;border-color:' + hoverColor + ';} QPushButton:hover{border:2px solid;border-color:' + selectedColor + ';background-color:' + buttonBackgroundColor + '}'
design_largeButtonTransparent = 'QPushButton{color:' + mainFontColor + ';font: ' + superFontSize + ' ' + headerFont + ';background-color: transparent;border-radius:0px;border:1px solid gray;padding-left: 10px; padding-right: 10px} QPushButton:focus{outline:0;border:1px solid;border-color:' + mainFontColor + ';} QPushButton:hover{border:2px solid;border-color:' + hoverColor + ';background-color:' + buttonBackgroundColor + '} QPushButton:pressed{background-color:' + hoverColor + ';border-style: inset;} QPushButton:checked{border: 2px solid ' + selectedColor + '} QPushButton:checked:pressed{background-color:' + selectedColor + ';border-style: inset;}'

design_editBox = 'QLineEdit{border: 1px solid gray;color:' + mainFontColor + ';background: transparent;selection-background-color:' + hoverColor + ';font: ' + smallFontSize + ' ' + monospaceFont + '} QLineEdit:focus{border: 1px solid ' + hoverColor + ';background:' + buttonBackgroundColor + '} QLineEdit:read-only{background: transparent; border: 1px transparent}'
design_editBoxTwo = 'QLineEdit{border-width: 0px 0px 1px 0px;border-style: solid;border-color: transparent transparent gray transparent;color: #cccccc;background: transparent;font: ' + smallFontSize + ' ' + monospaceFont + ';qproperty-alignment:AlignTop;padding-left:20px} QLineEdit:read-only{' + border + ';qproperty-alignment:AlignBottom;outline: 1px} QLineEdit:focus{border-color: transparent transparent white transparent;color:' + mainFontColor + '}'

design_GroupBox = 'QGroupBox{border-radius: 2px;background-color: ' + secondaryBackgroundColor + ';} QGroupBox::title{subcontrol-origin: margin;left: 10px;color:' + mainFontColor + ';}'
design_GroupBoxTwo = 'QGroupBox{border-radius: 0px;background-color: transparent;border: 1px solid grey} QGroupBox::title{subcontrol-origin: margin;left: 10px;color:' + mainFontColor + ';}'

design_comboBox = 'QComboBox{outline: 0; color:' + mainFontColor + ';background-color:transparent;border: 1px solid gray;padding: 0px 10px 0px 3px;font: ' + smallFontSize + ' ' + monospaceFont + '} QComboBox:!enabled{color:gray} QComboBox:selected{border: 1px solid ' + hoverColor + '} QComboBox:focus{border: 1px solid ' + mainFontColor + '}  QComboBox:on{border: 2px solid ' + hoverColor + '; background: ' + buttonBackgroundColor + '} QComboBox:drop-down{border: 1px solid white; color: ' + mainFontColor + '} QComboBox:drop-down:hover{border: 2px solid ' + hoverColor + '; background: ' + buttonBackgroundColor + '} QComboBox:drop-down:on{border: 2px solid ' + hoverColor + '; background: ' + buttonBackgroundColor + '} QAbstractItemView{color: ' + mainFontColor + ';background-color:' + mainBackgroundColor + ';selection-background-color:' + hoverColor + '} QListView{border: 2px solid ' + hoverColor + ';background-color:' + buttonBackgroundColor + '} QListView:item:hover{color:red}'

design_textEdit = 'QTextEdit{background-color: ' + mainBackgroundColor + ';font: ' + smallFontSize + ' ' + monospaceFont + ';color: ' + mainFontColor + ';border: 1px solid gray;selection-background-color:' + hoverColor + ';} QTextEdit:focus{border:1px solid ' + mainFontColor + '}'

# QComboBox:item{color:' + mainFontColor + ';border: 0px solid;} QListView{border: 2px solid ' + hoverColor + '; color:' + mainFontColor + ';background-color: ' + mainBackgroundColor + '} QListView:item:hover{background: ' + selectedColor + '; border: 2px solid red}

design_checkboxCross = 'QCheckBox{outline: 0} QCheckBox:indicator{width: 15px;height: 15px;background-color:' + buttonBackgroundColor + '} QCheckBox::indicator:unchecked:hover{image: url(./icons/checkboxUncheckedHover.png);border: 1px solid ' + hoverColor + '} QCheckBox::indicator:unchecked:pressed{image: url(./icons/checkboxUncheckedHover.png);border: 2px solid ' + hoverColor + '} QCheckBox:indicator:checked{image: url(./icons/checkboxChecked.png);} QCheckBox:indicator:checked:hover{image: url(./icons/checkboxChecked.png);border: 1px solid ' + selectedColor + '} QCheckBox:indicator:checked:pressed{image: url(./icons/checkboxChecked.png);border: 2px solid ' + selectedColor + '}'
design_checkboxTick = 'QCheckBox{outline: 0} QCheckBox:indicator{width: 15px;height: 15px;background-color:' + buttonBackgroundColor + '} QCheckBox::indicator:unchecked:hover{image: url(./icons/checkbox2UncheckedHover.png);border: 1px solid ' + hoverColor + '} QCheckBox::indicator:unchecked:pressed{image: url(./icons/checkbox2UncheckedHover.png);border: 2px solid ' + hoverColor + '} QCheckBox:indicator:checked{image: url(./icons/checkbox2Checked.png);} QCheckBox:indicator:checked:hover{image: url(./icons/checkbox2Checked.png);border: 1px solid ' + selectedColor + '} QCheckBox:indicator:checked:pressed{image: url(./icons/checkbox2Checked.png);border: 2px solid ' + selectedColor + '}'

design_calendar = 'QCalendarWidget QWidget #qt_calendar_navigationbar {background-color: ' + buttonBackgroundColor + ';color:' + mainFontColor + '} QCalendarWidget QToolButton {height: 20px;width: 100px;color:' + mainFontColor + ' ;icon-size:30px, 30px;background-color:' + buttonBackgroundColor + ';} QCalendarWidget QWidget #qt_calendar_prevmonth {width:40px; icon-size:30px, 30px} QCalendarWidget QWidget #qt_calendar_nextmonth {width:40px; icon-size:30px, 30px} QCalendarWidget QToolButton:hover {background-color:' + secondaryBackgroundColor + ';border:1px solid ' + hoverColor + '}  QCalendarWidget QToolButton:pressed{color:' + mainFontColor + ';background-color:' + hoverColor + '} QCalendarWidget QToolButton:menu-indicator{image:none} ' + design_tableMenu.replace('QMenu', 'QCalendarWidget QMenu') + ' QCalendarWidget QSpinBox {font-size:12px;background-color:' + mainBackgroundColor + ';color:' + mainFontColor + '; width: 75px; qproperty-alignment:AlignCenter} QCalendarWidget QSpinBox:up-button { subcontrol-origin: border;  subcontrol-position: top right;  width:25px; } QCalendarWidget QSpinBox:down-button {subcontrol-origin: border; subcontrol-position: bottom right;  width:25px;} QCalendarWidget QSpinBox:up-arrow {width:5px;  height:5px;} QSpinBox:down-arrow {width:5px;  height:5px;} QCalendarWidget QWidget {alternate-background-color: ' + secondaryBackgroundColor + '} QCalendarWidget QAbstractItemView{outline: 0px} QCalendarWidget QAbstractItemView:enabled{border:1px solid ' + hoverColor + ';font-size:12px;color:' + mainFontColor + ';background-color:' + buttonBackgroundColor + ';} QCalendarWidget QAbstractItemView:disabled{color: gray} QAbstractItemView:item:hover{color: ' + selectedColor + ';border: 1px solid ' + selectedColor + ';background-color:' + secondaryBackgroundColor + '} QAbstractItemView:item:selected{color: ' + mainFontColor + ';background-color:' + selectedColor + '}' 

design_inputDialog = 'QWidget {background-color: red;}' # testing

# SQL dictionaries
db_clientInfo_name = 'HHH_EF_Client_Info'
dict_toSQL_HHH_EF_Client_Info = {
    'client_number':'EFClientNumber', 
    'client_name':'ClientName', 
    'street_address':'ClientAddressStreet', 
    'suburb':'ClientAddressSuburb', 
    'state':'ClientAddressState', 
    'postcode':'ClientAddressPostcode', 
    'abn':'ClientABN', 
    'acn':'ClientACN', 
    'contact_phone':'ContactPhone', 
    'contact_email':'ContactEmail', 
    'contact_name':'ContactName'
}
dict_fromSQL_HHH_EF_Client_Info = dict((v, k) for k, v in dict_toSQL_HHH_EF_Client_Info.items()) #reverses dictionary mapping
client_number, client_name, street_address, suburb, state, postcode, abn, acn, contact_phone, contact_email, contact_name = list(dict_toSQL_HHH_EF_Client_Info.keys()) # assigns variables to use throughout code 

db_agreementInfo_name = 'HHH_EF_Agreement_Info'
dict_toSQL_HHH_EF_Agreement_Info = {
    'agreement_number':'EFAgreementNumber', 
    'mtx_number':'MtxBuyerNumber', 
    'bsb':'BankBSB', 
    'acc':'BankACC', 
    client_number:dict_toSQL_HHH_EF_Client_Info[client_number],
    'original_balance':'OriginalBalance', 
    'balloon_amount':'BalloonAmount', 
    'periodic_fee':'PeriodicFee', 
    'interest_rate':'InterestRatePerAnnum', 
    'periodic_repayment':'PeriodicRepayment', 
    'periods_per_year':'PeriodsPerAnnum', 
    'total_periods':'TotalPeriods', 
    'settlement_date':'SettlementDate', 
    'agreement_start_date':'AgreementStartDate', 
    'account_owner':'AccountOwner', 
    'agreement_status':'AgreementStatus', 
    'agreement_notes_asset':'AssetNotes', 
    'agreement_notes_account':'AccountNotes', 
    'agreement_notes_misc':'MiscellaneousNotes', 
}
dict_fromSQL_HHH_EF_Agreement_Info = dict((v, k) for k, v in dict_toSQL_HHH_EF_Agreement_Info.items()) #reverses dictionary mapping
agreement_number, mtx_number, bsb, acc, client_number, original_balance, balloon_amount, periodic_fee, interest_rate, periodic_repayment, periods_per_year, total_periods, settlement_date, agreement_start_date, account_owner, agreement_status, agreement_notes_asset, agreement_notes_account, agreement_notes_misc = list(dict_toSQL_HHH_EF_Agreement_Info.keys()) # assigns variables to use throughout modules 

db_payoutSchedule_name = 'HHH_EF_Schedule'
dict_toSQL_HHH_EF_Schedule = { # defining object names in code as well as assigning column names in SQL
    agreement_number:dict_toSQL_HHH_EF_Agreement_Info[agreement_number], 
    mtx_number:dict_toSQL_HHH_EF_Agreement_Info[mtx_number], 
    client_name:dict_toSQL_HHH_EF_Client_Info[client_name], 
    'payment_number':'EFPaymentNumber', 
    'repayment_date':'OriginalRepaymentDate', 
    'opening_balance':'OpeningBalance', 
    'interest_component':'InterestComponent', 
    'principal_component':'PrincipalComponent', 
    periodic_repayment:dict_toSQL_HHH_EF_Agreement_Info[periodic_repayment], 
    'closing_balance':'ClosingBalance', 
    'payment_status':'PaymentStatus', 
    'completion_date':'CompletionDate', 
    'payment_notes':'PaymentNotes'
}
dict_fromSQL_HHH_EF_Schedule = dict((v, k) for k, v in dict_toSQL_HHH_EF_Schedule.items()) #reverses dictionary mapping
agreement_number, mtx_number, client_name, payment_number, repayment_date, opening_balance, interest_component, principal_component, periodic_repayment, closing_balance, payment_status, completion_date, payment_notes  = list(dict_toSQL_HHH_EF_Schedule.keys()) # assigns variables to use throughout code 

db_subordinate_name = 'HHH_EF_Subordinates'
dict_toSQL_HHH_EF_Subordinates = { # defining object names in code as well as assigning column names in SQL - ORDER OF COLUMNS MUST BE SAME
    payment_number:dict_toSQL_HHH_EF_Schedule[payment_number], 
    'subordinate_number':'EFSubordinateNumber', 
    'subordinate_date':'SubordinateDate', 
    'subordinate_type':'ChargeType', 
    'subordinate_amount':'Amount', 
    'subordinate_gst':'GST', 
    'subordinate_status':'SubordinateStatus', 
    'subordinate_value_date':'SubordinateValueDate', 
    'subordinate_notes':'SubordinateNotes', 
    'subordinate_created_by':'CreatedBy'
}
dict_fromSQL_HHH_EF_Subordinates = dict((v, k) for k, v in dict_toSQL_HHH_EF_Subordinates.items()) #reverses dictionary mapping
payment_number, subordinate_number, subordinate_date, subordinate_type, subordinate_amount, subordinate_gst, subordinate_status, subordinate_value_date, subordinate_notes, subordinate_created_by = list(dict_toSQL_HHH_EF_Subordinates.keys()) # Maintain variable order with column order 

db_userAdmin_name = 'mPowerUserAdmin'
dict_toSQL_userAdmin = { # defining object names in code as well as assigning column names in SQL - ORDER OF COLUMNS MUST BE SAME
    'userAdmin_username':'UserName', 
    'userAdmin_password':'Password', 
    'userAdmin_latestLogin':'LatestLogin', 
    'userAdmin_latestPasswordUpdate':'LatestPasswordUpdate', 
    'userAdmin_firstName':'FirstName', 
    'userAdmin_lastName':'LastName', 
    'userAdmin_emailAddress':'EmailAddress', 
    'userAdmin_authority':'Authority'
}
dict_fromSQL_userAdmin = dict((v, k) for k, v in dict_toSQL_userAdmin.items()) #reverses dictionary mapping
userAdmin_username, userAdmin_password, userAdmin_latestLogin, userAdmin_latestPasswordUpdate, userAdmin_firstName, userAdmin_lastName, userAdmin_emailAddress, userAdmin_authority = list(dict_toSQL_userAdmin.keys()) # Maintain variable order with column order 

db_userCredentials_name = 'mPowerCredentials'
dict_toSQL_userCredentials = { # defining object names in code as well as assigning column names in SQL - ORDER OF COLUMNS MUST BE SAME
    'userCredentials_user':'mPowerUserName', 
    'userCredentials_application':'WebName', 
    'userCredentials_username':'CredentialUserName', 
    'userCredentials_password':'CredentialPW', 
    'userCredentials_credential1':'OtherCredential1', 
    'userCredentials_credential2':'OtherCredential2'
}
dict_fromSQL_userCredentials = dict((v, k) for k, v in dict_toSQL_userCredentials.items()) #reverses dictionary mapping
userCredentials_user, userCredentials_application, userCredentials_username, userCredentials_password, userCredentials_credential1, userCredentials_credential2 = list(dict_toSQL_userCredentials.keys()) # Maintain variable order with column order 

db_debtorReceipts_name = 'DebtorReceipts'
dict_toSQL_debtorReceipts = {
    'debtorReceipts_client':'Client', 
    'debtorReceipts_value':'DebtorReceipts', 
    'debtorReceipts_status':'AQStatus', 
    'debtorReceipts_date':'Date'
}
dict_fromSQL_debtorReceipts = dict((v, k) for k, v in dict_toSQL_debtorReceipts.items()) #reverses dictionary mapping
debtorReceipts_client, debtorReceipts_value, debtorReceipts_status, debtorReceipts_date = list(dict_toSQL_debtorReceipts.keys()) # Maintain variable order with column order 

db_EFInstructions_name = 'mPowerEFInstructions'
dict_toSQL_EFInstructions = {
    'EFInstructions_valueDate':'ValueDate', 
    'EFInstructions_agreementNumber':'AgreementID', 
    'EFInstructions_transactionType':'TransactionType', 
    'EFInstructions_status':'Status', 
    'EFInstructions_amount':'Amount'
}
dict_fromSQL_EFInstructions = dict((v, k) for k, v in dict_toSQL_EFInstructions.items()) #reverses dictionary mapping
EFInstructions_valueDate, EFInstructions_agreementNumber, EFInstructions_transactionType, EFInstructions_status, EFInstructions_amount = list(dict_toSQL_EFInstructions.keys()) # Maintain variable order with column order 

df_DFClientInfo_name = 'DFClientInfo'
dict_toSQL_DFClientInfo = {
    'DFClientInfo_AQClientNumber':'AQClient#', 
    'DFClientInfo_AQClientName':'AQClientName', 
    'DFClientInfo_ClientABN':'ClientABN', 
    'DFClientInfo_MTXBuyerAccount':'MTXBuyerAccount', 
    'DFClientInfo_MTXSellerAccount':'MTXSellerAccount', 
    'DFClientInfo_Address1':'ClientAddress1', 
    'DFClientInfo_Address2':'ClientAddress2', 
    'DFClientInfo_Address3':'ClientAddress3', 
    'DFClientInfo_Address4':'ClientAddress4', 
    'DFClientInfo_ContactName1':'ContactName1', 
    'DFClientInfo_ClientEmail1':'ClientEmail1', 
    'DFClientInfo_ContactName2':'ContactName2', 
    'DFClientInfo_ClientEmail2':'ClientEmail2', 
    'DFClientInfo_ContactName3':'ContactName3', 
    'DFClientInfo_ClientEmail3':'ClientEmail3', 
    'DFClientInfo_EntityName':'EntityName', 
    'DFClientInfo_EntityType':'EntityType', 
    'DFClientInfo_ClientACN':'ClientACN', 
}
dict_fromSQL_DFClientInfo = dict((v, k) for k, v in dict_toSQL_DFClientInfo.items()) #reverses dictionary mapping
DFClientInfo_AQClientNumber, DFClientInfo_AQClientName, DFClientInfo_ClientABN, DFClientInfo_MTXBuyerAccount, DFClientInfo_MTXSellerAccount, DFClientInfo_Address1, DFClientInfo_Address2, DFClientInfo_Address3, DFClientInfo_Address4, DFClientInfo_ContactName1, DFClientInfo_ClientEmail1, DFClientInfo_ContactName2, DFClientInfo_ClientEmail2, DFClientInfo_ContactName3, DFClientInfo_ClientEmail3, DFClientInfo_EntityName, DFClientInfo_EntityType, DFClientInfo_ClientACN = list(dict_toSQL_DFClientInfo.keys()) # Maintain variable order with column order 

db_Events_name = 'HHH_Events'
dict_toSQL_Events = {
    # 'events_ID':'EventID', # Auto-Increments by 1 on every new row - exclude from dictionary as you can't insert explicit values for auto-incrementing columns 
    'events_EventTime':'EventTime', 
    'events_UserName':'UserName', 
    'events_Event':'Event', 
    'events_EventDescription':'EventDescription',
}
dict_fromSQL_Events = dict((v, k) for k, v in dict_toSQL_Events.items()) #reverses dictionary mapping
events_EventTime, events_UserName, events_Event, events_EventDescription = list(dict_toSQL_Events.keys()) # Maintain variable order with column order 

dict_eventCategories = { # record of event categories that can be written into the events database for audit trail records
    'error': 'Critical Error',
    'settings': 'Settings',
    'segments': 'Segments',
    'debtorReceipts': 'Debtor Receipts',
    'halo': 'Halo Login',
    'b2b': 'B2B Agreements',
    'equipmentFinance': 'Equipment Finance'
}

dict_agreementStatus = {
    'agreementStatus_active':'Active',  
    'agreementStatus_closed':'Closed'
}
agreementStatus_active, agreementStatus_closed = list(dict_agreementStatus.keys()) # assigns variables to use throughout code 

dict_paymentStatus = {
    'paymentStatus_pending':'Pending', 
    'paymentStatus_due':'Due', 
    'paymentStatus_uncleared':'Uncleared', 
    'paymentStatus_overdue':'Overdue', 
    'paymentStatus_rescheduled':'Rescheduled', 
    'paymentStatus_received':'Received', 
    'paymentStatus_credited':'Credited'
}
paymentStatus_pending, paymentStatus_due, paymentStatus_uncleared, paymentStatus_overdue, paymentStatus_rescheduled, paymentStatus_received, paymentStatus_credited = list(dict_paymentStatus.keys()) # assigns variables to use throughout code 
defaultEFPaymentsTableMenuFilter = [dict_paymentStatus[paymentStatus_overdue], dict_paymentStatus[paymentStatus_due], dict_paymentStatus[paymentStatus_pending], dict_paymentStatus[paymentStatus_rescheduled]] # Default selected payment statuses for EF Subordinate table
outstandingStatuses = [dict_paymentStatus[paymentStatus_overdue], dict_paymentStatus[paymentStatus_due], dict_paymentStatus[paymentStatus_pending], dict_paymentStatus[paymentStatus_rescheduled], dict_paymentStatus[paymentStatus_uncleared]]

dict_subordinateType = {
    'subordinateType_repayment':'Periodic Repayment', 
    'subordinateType_dishonour':'Dishonour Fee', 
    'subordinateType_reschedule':'Reschedule Fee', 
    'subordinateType_periodic':'Periodic Fee', 
    'subordinateType_miscellaneous':'Miscellaneous', 
}
subordinateType_repayment, subordinateType_dishonour, subordinateType_reschedule, subordinateType_periodic, subordinateType_miscellaneous = list(dict_subordinateType.keys()) # assigns variables to use throughout code

list_subordinateGST = [dict_subordinateType[subordinateType_periodic], dict_subordinateType[subordinateType_dishonour], dict_subordinateType[subordinateType_reschedule]] # used for determining if a subordinate carries gst charge

dict_statementStyles = {
    'styleHeading':{'font':'Roboto Condensed', 'bold':True, 'size':20, 'color':'1d396a', 'horizontal-alignment':'left', 'vertical-alignment':'center'}, 
    'styleHeader':{'font':'Roboto Condensed', 'bold':False, 'size':14, 'color':'58595b', 'background-color':'d1d3d4', 'vertical-alignment':'center', 'horizontal-alignment':'left'}, 
    'styleLabel':{'font':'Univers', 'bold':False, 'size':11, 'color':'58595b', 'vertical-alignment':'center', 'horizontal-alignment':'left'}, 
    'styleValue':{'font':'Univers', 'bold':False, 'size':11, 'color':'000000', 'vertical-alignment':'center', 'horizontal-alignment':'right'}
}
styleHeading, styleHeader, styleLabel, styleValue = list(dict_statementStyles.keys())


# Other dictionaries
dict_permissions = {
    'HHH_maximum': ['changeUserAuthorities', 'changeHaloCredentials', 'viewHaloPassword', 'sendSegments', 'doDebtorReceipts', 'commitB2BTransactions', 'doMailMerges', 'loginToHalo', 'loginToHaloDemo', 'EFAddAgreement', 'EFViewAgreement', 'EFViewPayment', 'EFUpdatePayment', 'EFViewSubordinate', 'EFAddSubordinate', 'EFUpdateSubordinate','EFMaintainAgreement'],
    'HHH_default': ['loginToHaloDemo', 'EFViewAgreement', 'EFViewPayment', 'EFViewSubordinate'],
    'HHH_accountManager': ['sendSegments', 'doDebtorReceipts', 'commitB2BTransactions', 'doMailMerges', 'loginToHalo', 'loginToHaloDemo', 'EFAddAgreement', 'EFViewAgreement', 'EFViewPayment', 'EFUpdatePayment', 'EFViewSubordinate', 'EFAddSubordinate', 'EFUpdateSubordinate','EFMaintainAgreement'],
}
dict_b2bSameDayPaymentFees = {

}
dict_b2bTransactions = {
    'Same Day Payment':{
        'includeGST':False, 
        'AQInstructions':'RTGS', 
        'MTXInstructions':'debit'
    }, 
    'Overnight Payment':{
        'includeGST':False, 
        'AQInstructions':None, 
        'MTXInstructions':'debit'
    }, 
    'Transfer to Reserve':{
        'includeGST':False, 
        'AQInstructions':None, 
        'MTXInstructions':'debit'
    }, 
    'Reserve Reduction':{
        'includeGST':False, 
        'AQInstructions':None, 
        'MTXInstructions':'credit'
    }, 
    'Field Review Fee':{
        'includeGST':True, 
        'AQInstructions':None, 
        'MTXInstructions':'debit'
    }, 
    'Debit Note':{
        'includeGST':False, 
        'AQInstructions':None, 
        'MTXInstructions':'debit'
    }, 
    'Credit Note':{
        'includeGST':False, 
        'AQInstructions':None, 
        'MTXInstructions':'credit'
    }
}

class widgetEffects():
    # Makes window translucent and click-through
    @staticmethod
    def translucifyWindow(window, state, geometry):
        window.setWindowFlag(QtCore.Qt.WindowTransparentForInput, state)
        window.setWindowOpacity(translucifyOpacityLevel) if state is True else window.setWindowOpacity(1)
        window.show() # necessary to navigate around known bug when setting opacity

    # Dsiplays message at defined label before fading after certain time
    @staticmethod
    def flashMessage(label, duration, message, messageType):
        dict_messageStyles = {
            'normal':design_textSmall, 
            'warning':design_textSmallWarning, 
            'error':design_textSmallError, 
            'success':design_textSmallSuccess
        }
        label.setStyleSheet(dict_messageStyles[messageType])
        label.setText(message)
        effect = QtWidgets.QGraphicsOpacityEffect(label)
        label.setGraphicsEffect(effect)
        label.fade = QtCore.QPropertyAnimation(effect, b'opacity')
        label.fade.setDuration(duration)
        label.fade.setStartValue(1)
        label.fade.setEndValue(0)
        label.fade.setEasingCurve(QtCore.QEasingCurve.InExpo)
        label.fade.start()

    @staticmethod
    def animateWindowOffScreen(window, state, geometry):
        screenRes = QtWidgets.QApplication.desktop().screenGeometry()
        screenW, screenH = screenRes.width(), screenRes.height()
        window.anim = QtCore.QPropertyAnimation(window, b'geometry')
        window.anim.setDuration(1000)
        if state is True:
            window.anim.setStartValue(QtCore.QRect(geometry.x(), geometry.y(), geometry.width(), geometry.height()))
            window.anim.setEndValue(QtCore.QRect(screenW-geometry.width()-50, screenH-50, geometry.width(), geometry.height()))
        else:
            window.anim.setStartValue(QtCore.QRect(geometry.x(), screenH-50, geometry.width(), geometry.height()))
            window.anim.setEndValue(QtCore.QRect(geometry.x(), screenH-geometry.height()-50, geometry.width(), geometry.height()))
        window.anim.setEasingCurve(QtCore.QEasingCurve.OutElastic)
        window.anim.start()

class HHHWindowWidget(QtWidgets.QWidget):

    def __init__(self, username, geometry, padding={'left':0, 'top':20, 'right':0, 'bottom':20}, title='HHHchildWindow', updateButton=None, parent=None):
        super().__init__(parent)
        QtWidgets.QApplication.setEffectEnabled(QtCore.Qt.UI_AnimateCombo, False) # Disables ComboBox drop-down animation
        self.username = username
        self.padding = padding
        self.updateButton = updateButton
        self.setWindowFlags(QtCore.Qt.FramelessWindowHint | QtCore.Qt.Tool | QtCore.Qt.WindowStaysOnTopHint)
        self.setGeometry(geometry.x(), geometry.y(), geometry.width(), geometry.height())
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground)                 
        self.title = title
        self.loadWindow()          

    def paintEvent(self, event):
        s = self.size()
        qp = QtGui.QPainter()
        qp.begin(self)
        qp.setRenderHint(QtGui.QPainter.Antialiasing, True)
        qp.setPen(QtGui.QColor(selectedColor))
        qp.setBrush(QtGui.QColor(0, 0, 0, 150))
        qp.drawRect(0, 0, s.width(), s.height())
        qp.end()

    def keyPressEvent(self, event):
        if event.type() == QtCore.QEvent.KeyPress:
            if event.key() == QtCore.Qt.Key_Escape:
                self.close()
        return super(HHHWindowWidget, self).keyPressEvent(event)
    
    def loadWindow(self):
        # total window, including transparent padding
        vBox_window = QtWidgets.QVBoxLayout()
        vBox_window.setContentsMargins(self.padding['left'], self.padding['top'], self.padding['right'], self.padding['bottom'])
        vBox_window.setSpacing(0)
        vBox_window.setAlignment(QtCore.Qt.AlignTop) 
        self.setLayout(vBox_window)

        self.widgetFrame = QtWidgets.QWidget()         
        self.widgetFrame.setWindowTitle(self.title)
        self.widgetFrame.setMinimumSize(self.geometry().width() - self.padding['left'] - self.padding['right'], 10)
        self.widgetFrame.setObjectName(windowObject)
        self.widgetFrame.setStyleSheet('QWidget#' +  windowObject  + ' {' + design_background + ';border: 2px solid ' + selectedColor + '}')
        vBox_window.addWidget(self.widgetFrame)

        # Main vBox inside window frame
        vBox = QtWidgets.QVBoxLayout()
        vBox.setContentsMargins(10, 10, 10, 10)
        vBox.setSpacing(5)
        vBox.setAlignment(QtCore.Qt.AlignTop) 
        self.widgetFrame.setLayout(vBox)

        self.vBox_contents = QtWidgets.QVBoxLayout()
        self.vBox_contents.setContentsMargins(0, 0, 0, 0)
        self.vBox_contents.setSpacing(10)
        self.vBox_contents.setAlignment(QtCore.Qt.AlignTop) 
        vBox.addLayout(self.vBox_contents)

        # vBox.addStretch(1)

        # Bottom Buttons
        hBox_bottomButtons = QtWidgets.QHBoxLayout()
        hBox_bottomButtons.setContentsMargins(0, 0, 0, 0)
        hBox_bottomButtons.setSpacing(5)
        hBox_bottomButtons.setAlignment(QtCore.Qt.AlignTop) 
        vBox.addLayout(hBox_bottomButtons)        

        if self.updateButton:
            self.button_update = QtWidgets.QPushButton(str(self.updateButton), default=True)
            self.button_update.setFixedHeight(20)
            self.button_update.setStyleSheet(design_smallButtonTransparent)
            hBox_bottomButtons.addWidget(self.button_update)

        self.button_cancel = QtWidgets.QPushButton('Cancel', clicked=self.close, default=True)
        self.button_cancel.setFixedHeight(20)
        self.button_cancel.setStyleSheet(design_smallButtonTransparent)
        hBox_bottomButtons.addWidget(self.button_cancel)

        # Bottom Status
        hBox_status = QtWidgets.QHBoxLayout()
        hBox_status.setContentsMargins(0, 0, 0, 0)
        hBox_status.setSpacing(0)
        hBox_status.setAlignment(QtCore.Qt.AlignTop) 
        vBox.addLayout(hBox_status)

        self.heading_status = QtWidgets.QLabel('')
        self.heading_status.setStyleSheet(design_textSmall)
        vBox.addWidget(self.heading_status)

class dragDropButton(QtWidgets.QPushButton):
    def __init__(self, parent=None, objectName=None, text=None, dragText=None, droppedConnect=None):
        super().__init__(parent)
        self.buttonText = text
        self.dragText = dragText   
        self.droppedConnect = droppedConnect    
        self.setAcceptDrops(True) 
        self.setStyleSheet(design_smallButtonTransparentDrag)
        self.setFocusPolicy(QtCore.Qt.NoFocus)
        self.setObjectName(objectName)
        self.setText(self.buttonText)    

    def dragEnterEvent(self, event):
        self.setStyleSheet(design_smallButtonTransparentDrag)
        self.buttonText = self.text() # Overrides default text with latest button text
        if event.mimeData().hasUrls():
            self.setText(self.dragText)
            self.setStyleSheet(design_smallButtonTransparentDragEnter)
            event.acceptProposedAction()
        else:
            super().dragEnterEvent(event)

    def dragLeaveEvent(self, event):
        self.setText(self.buttonText)
        self.setStyleSheet(design_smallButtonTransparentDrag) 
        event.accept()

    def dragMoveEvent(self, event):
        super().dragMoveEvent(event)

    def dropEvent(self, event):
        self.setText(self.buttonText)
        self.setStyleSheet(design_smallButtonTransparentDrag)
        if self.droppedConnect is not None:
            self.droppedConnect(event, self.objectName())

class tableComboBox(QtWidgets.QComboBox):
    def __init__(self, scrollWidget=None, objectName=None, *args, **kwargs):
        super(tableComboBox, self).__init__(*args, **kwargs)  
        self.scrollWidget=scrollWidget
        self.setFocusPolicy(QtCore.Qt.NoFocus)
        if objectName is not None:
            self.setObjectName(objectName)

    def wheelEvent(self, *args, **kwargs):
        if self.hasFocus():
            return QtWidgets.QComboBox.wheelEvent(self, *args, **kwargs)
        else:
            return self.scrollWidget.wheelEvent(*args, **kwargs)

class CustomMenu(QtWidgets.QMenu):

    signal_checkableAction = QtCore.pyqtSignal(object)

    def __init__(self, menuItems=[], parent=None):
        super().__init__()
        self.setStyleSheet(design_tableMenu)
        self.addMenuItems(self, menuItems)
        self.parent = parent 
        self.aboutToShow.connect(self.highlightParentRow)

    def addMenuItems(self, menu, menuItems):
        menu.installEventFilter(self)
        for item in menuItems:
            if item['type'] == 'action':
                item['object'] = QtWidgets.QAction(item['displayText'], menu, checkable=item['checkable'], checked=item['checkedState'])
                menu.addAction(item['object'])
            elif item['type'] == 'section':
                menu.addSeparator()
            elif item['type'] == 'widgetAction':
                pass
            elif item['type'] == 'menu':
                self.subMenu = menu.addMenu(item['displayText'])
                self.subMenu.setStyleSheet(design_tableMenu)
                self.addMenuItems(self.subMenu, item['subMenu'])
            else:
                print('Menu item not defined: ' + item['type'])

    def eventFilter(self, obj, event): # Event filter for checkable items (must define SLOT to receive signal with)
        if event.type() == QtCore.QEvent.MouseButtonRelease:
            if obj.activeAction():
                if obj.activeAction().isCheckable():  # if selected action is checkable
                    obj.activeAction().trigger() # eat hideEvent but trigger the function
                    self.checkableItemTriggered(obj.activeAction())
                    return True
        return super(CustomMenu, self).eventFilter(obj, event)

    def checkableItemTriggered(self, sender):
        self.signal_checkableAction.emit(sender)
    
    def highlightParentRow(self):
        if self.parent:
            self.parent.selectRow(self.parent.currentIndex().row())

class HHHTableWidget(QtWidgets.QTableView):

    def __init__(self, dataFrame=None, defaultSectionSize=50, fixedHeight=None, cornerWidget=None, contextMenu=None, doubleClicked=None, *args, **kwargs):
        super(HHHTableWidget, self).__init__(*args, **kwargs) 
        QtWidgets.QApplication.setEffectEnabled(QtCore.Qt.UI_AnimateTooltip, False) # Disables tooltip animation
        self.setStyleSheet(design_table)
        self.setFocusPolicy(QtCore.Qt.StrongFocus)      
        self.setSizeAdjustPolicy(QtWidgets.QAbstractScrollArea.AdjustToContents)
        if dataFrame is not None:
            self.setModel(dataFrame)
        
        self.horizontalScrollBar().setStyleSheet(design_tableHorizontalScrollBar)
        self.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAsNeeded)
        self.setHorizontalScrollMode(QtWidgets.QAbstractItemView.ScrollPerPixel)
        self.setWordWrap(True)  
        self.defaultSectionSize = defaultSectionSize
        if fixedHeight:
            self.setMinimumHeight(fixedHeight)           
            self.setMaximumHeight(fixedHeight)

        if doubleClicked:
            self.doubleClicked.connect(doubleClicked)
        # self.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.ResizeToContents)
        # for i in range(self.columnCount()):
        #     self.horizontalHeaderItem(i).setTextAlignment(QtCore.Qt.AlignLeft)
        self.horizontalHeader().setHighlightSections(0)
        self.horizontalHeader().setStyleSheet(design_tableHeader)
        self.horizontalHeader().setMinimumSectionSize(10)
        self.horizontalHeader().setDefaultSectionSize(defaultSectionSize)
        self.horizontalHeader().setDefaultAlignment(QtCore.Qt.AlignLeft)
        # self.horizontalHeader().setSectionsMovable(True)
            
        self.verticalHeader().setVisible(False)
        self.verticalScrollBar().setStyleSheet(design_tableVerticalScrollBar)
        self.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOn)
        self.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.setWordWrap(False)
        if cornerWidget:
            self.installCornerWidget(cornerWidget)
        self.lastSelectedRow = None
        self.installEventFilter(self)
        self.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)        
        self.customContextMenuRequested.connect(self.displayContextMenu)
        self.contextMenu = contextMenu
        self.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)

    def eventFilter(self, obj, event):
        if event.type() != QtCore.QEvent.KeyPress:
            return super(HHHTableWidget, self).eventFilter(obj, event)
        if event.matches(QtGui.QKeySequence.Copy):
            self.copyTable()
            return True
        elif event.key() in [QtCore.Qt.Key_Return, QtCore.Qt.Key_Enter]:
            self.doubleClicked.emit(self.currentIndex())
        elif event.key() == QtCore.Qt.Key_Tab:
            self.focusOutEvent(QtGui.QFocusEvent(QtCore.QEvent.FocusOut, QtCore.Qt.TabFocusReason))
            event.ignore()
        elif event.key() == QtCore.Qt.Key_Backtab:
            self.focusOutEvent(QtGui.QFocusEvent(QtCore.QEvent.FocusOut, QtCore.Qt.BacktabFocusReason))
            event.ignore()
        elif ((event.modifiers() == QtCore.Qt.ControlModifier) and (event.key() == QtCore.Qt.Key_Up)) or event.key() == QtCore.Qt.Key_Home:
            self.scrollToTop()
            self.selectRow(0)
        elif ((event.modifiers() == QtCore.Qt.ControlModifier) and (event.key() == QtCore.Qt.Key_Down)) or event.key() == QtCore.Qt.Key_End:
            self.scrollToBottom()
            self.selectRow(self.model().rowCount()-1)
        return super(HHHTableWidget, self).eventFilter(obj, event)  

    def focusInEvent(self, event):
        if self.lastSelectedRow:
            self.setCurrentIndex(self.model().index(self.lastSelectedRow, 1))
        else:
            try:
                self.setCurrentIndex(self.model().index(0, 1))
            except:
                pass
        return super(HHHTableWidget, self).focusInEvent(event)

    def focusOutEvent(self, event):
        self.lastSelectedRow = self.currentIndex().row()
        if event.reason() != QtCore.Qt.PopupFocusReason:
            self.clearSelection()                
        lastVisibleItem = self.indexAt(QtCore.QPoint(0, 0))
        if lastVisibleItem is not None:
            self.scrollTo(lastVisibleItem, QtWidgets.QAbstractItemView.PositionAtTop)
        return super(HHHTableWidget, self).focusOutEvent(event)    

    def setModel(self, dataFrame):
        model = self.pandasModel(dataFrame)
        proxyModel = QtCore.QSortFilterProxyModel()
        proxyModel.setSourceModel(model)
        super(HHHTableWidget, self).setModel(proxyModel)
        self.horizontalHeader().setStretchLastSection(False)     
        self.horizontalHeader().resizeSections(QtWidgets.QHeaderView.ResizeToContents)
        self.horizontalHeader().setStretchLastSection(True)        

    def displayContextMenu(self, pos):
        if self.contextMenu:
            self.contextMenu(pos)   

    def copyTable(self): # copies current table to clipboard
        # If user presses copy widget instead of Ctrl + C, it is possible no row is currently selected
        try: 
            currentRow = self.currentRow()
        except:
            currentRow = None
        self.setSelectionMode(QtWidgets.QAbstractItemView.ExtendedSelection)
        self.selectAll()
        data = self.selectedIndexes()
        self.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self.setSelectionBehavior(QtWidgets.QTableWidget.SelectRows)  
        if currentRow:
            self.selectRow(currentRow)
        else:
            self.clearSelection()      
        if data:
            rows = max([index.row() for index in data]) + 1 
            cols = max([index.column() for index in data]) + 1
            table = [[''] * cols for _ in range(rows)] # empty table list
            for index in data:
                table[index.row()][index.column()] = index.data()
            # table headers
            headerItems = [self.model().headerData(i, QtCore.Qt.Horizontal, role=QtCore.Qt.DisplayRole) for i in range(cols)]
            table.insert(0, headerItems)
            QtWidgets.QApplication.clipboard().setText('\n'.join(['\t'.join([str(x) for x in tableRow]) for tableRow in table]))

    def installCornerWidget(self, cornerWidget):
        cornerButton = QtWidgets.QPushButton()
        cornerButton.setStyleSheet(design_smallButtonTransparent.replace(largeFontSize, smallFontSize))         
        # copies current table to clipboard
        if cornerWidget == 'copy':
            cornerButton.clicked.connect(self.copyTable)
            buttonText = ''
            buttonTool = 'Copy table to clipboard'
        cornerButton.setFocusPolicy(QtCore.Qt.NoFocus)
        cornerButton.setText(buttonText)
        cornerButton.setToolTip(buttonTool)
        self.setCornerWidget(cornerButton)

    class pandasModel(QtCore.QAbstractTableModel):
        def __init__(self, data, parent=None):
            QtCore.QAbstractTableModel.__init__(self, parent)
            self._data = data

        def rowCount(self, parent=None):
            return len(self._data.values)

        def columnCount(self, parent=None):
            return self._data.columns.size

        def data(self, index, role=QtCore.Qt.DisplayRole):
            if index.isValid():
                if role == QtCore.Qt.DisplayRole:
                    return str(self._data.values[index.row()][index.column()])
            return None

        def headerData(self, col, orientation, role):
            if orientation == QtCore.Qt.Horizontal and role == QtCore.Qt.DisplayRole:
                return self._data.columns[col]
            return None

class calendarLineEditWidget(QtWidgets.QLineEdit):

    def __init__(self, parent, objName=None, fixedWidth=None, calWidth=270, calHeight=300, adjustmentWidget=None, paddingLeft=None, positionTweakX=0, positionTweakY=0, *args, **kwargs):
        super(calendarLineEditWidget, self).__init__(*args, **kwargs)  
        self.parent = parent # parent of calendar widget - should be parent of lineedit
        self.calWidth = calWidth       
        self.calHeight = calHeight
        self.adjustmentWidget = adjustmentWidget # If calendar gets cutoff, set parent to be the window, and this to be the parent of lineedit
        self.positionTweakX = positionTweakX
        self.positionTweakY = positionTweakY
        self.setPlaceholderText(dateFormat)
        self.setValidator(QtGui.QRegExpValidator(QtCore.QRegExp('^\d\d/\d\d/\d\d\d\d$')))
        self.textChanged.connect(self.dateChanged)
        
        if objName is not None:
            self.objName = str(objName)
            self.setObjectName('edit_' + self.objName)
        if fixedWidth is not None:
            self.setFixedWidth(int(fixedWidth))
        if paddingLeft is not None:
            self.setStyleSheet(design_editBoxTwo.replace('20px', str(paddingLeft)))
        else:
            self.setStyleSheet(design_editBoxTwo)
    
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

    def mouseReleaseEvent(self, event):
        try:
            self.childWindow.setVisible(not self.childWindow.isVisible())
        except:
            self.openCalWidget(parent=self.parent, objName=self.objName, geometry=self.geometry(), calWidth=self.calWidth, calHeight=self.calHeight, adjustmentWidget=self.adjustmentWidget, positionTweakX=self.positionTweakX, positionTweakY=self.positionTweakY)

    def focusInEvent(self, event):
        if event.reason() in [QtCore.Qt.TabFocusReason, QtCore.Qt.BacktabFocusReason]:
            try:
                self.childWindow.setVisible(True)
            except:
                self.openCalWidget(parent=self.parent, objName=self.objName, geometry=self.geometry(), calWidth=self.calWidth, calHeight=self.calHeight, adjustmentWidget=self.adjustmentWidget, positionTweakX=self.positionTweakX, positionTweakY=self.positionTweakY)

    def focusOutEvent(self, event):
        try: # see if cal was clicked - if so, lineedit retains focus
            if QtWidgets.QApplication.focusWidget().parent() == self.childWindow:
                return
        except:
            pass
        # otherwise lose focus and close (hide) cal
        super(calendarLineEditWidget, self).focusOutEvent(event)
        try:
            self.closeCalWidget()
            # TODO if date text invalid, set blank
        except:
            pass

    def openCalWidget(self, parent, objName, geometry, calWidth, calHeight, adjustmentWidget, positionTweakX, positionTweakY):

        class customCalWidget(QtWidgets.QCalendarWidget):

            def __init__(self, parent):
                super(customCalWidget, self).__init__(parent)
                self.setObjectName('calendar_' + str(objName))
                self.setDateRange(QtCore.QDate(2000, 1, 1), QtCore.QDate(3000, 1, 1))
                self.setGridVisible(False)
                self.setFocusPolicy(QtCore.Qt.NoFocus)
                self.adjustmentWidget = adjustmentWidget
                adjustmentX = positionTweakX if self.adjustmentWidget is None else self.adjustmentWidget.x() + positionTweakX
                adjustmentY = positionTweakY + 5 if self.adjustmentWidget is None else self.adjustmentWidget.y() + positionTweakY + 5
                self.setGeometry(geometry.x()-calWidth + geometry.width() + adjustmentX, geometry.y() + geometry.height() + adjustmentY, calWidth, calHeight)
                self.setVerticalHeaderFormat(QtWidgets.QCalendarWidget.NoVerticalHeader)
                saturdayFormat = self.weekdayTextFormat(QtCore.Qt.Saturday)
                saturdayFormat.setForeground(QtGui.QBrush(QtCore.Qt.gray, QtCore.Qt.SolidPattern))
                self.setWeekdayTextFormat(QtCore.Qt.Saturday, saturdayFormat)
                sundayFormat = self.weekdayTextFormat(QtCore.Qt.Sunday)
                sundayFormat.setForeground(QtGui.QBrush(QtCore.Qt.gray, QtCore.Qt.SolidPattern))
                self.setWeekdayTextFormat(QtCore.Qt.Sunday, sundayFormat)
                self.findChild(QtWidgets.QWidget, 'qt_calendar_prevmonth').setIcon(QtGui.QIcon(icons_dir + '\\arrow_left.png'))
                self.findChild(QtWidgets.QWidget, 'qt_calendar_nextmonth').setIcon(QtGui.QIcon(icons_dir + '\\arrow_right.png'))
                # set delegate - necessary for hover styles over the abstractItemView items
                delegate = QtWidgets.QStyledItemDelegate(self)
                self.findChild(QtWidgets.QAbstractItemView).setItemDelegate(delegate)
                self.setStyleSheet(design_calendar) 
                # custom today() button
                todayButton = QtWidgets.QPushButton('Today', objectName='qt_calendar_today', clicked=self.gotoToday)
                todayButton.setStyleSheet(design_smallButtonTransparent.replace(buttonBackgroundColor, secondaryBackgroundColor).replace('transparent', buttonBackgroundColor).replace('1px', '0px').replace('2px', '1px'))
                self.findChild(QtWidgets.QVBoxLayout).insertWidget(1, todayButton)     

            def gotoToday(self):
                self.setSelectedDate(QtCore.QDate.currentDate())

        try:
            self.childWindow.setVisible(True)
        except:
            self.childWindow = customCalWidget(parent)
            self.childWindow.selectionChanged.connect(self.updateDate)
            self.childWindow.activated.connect(self.updateDate)   
            self.childWindow.activated.connect(self.setFocus)                  
            self.childWindow.activated.connect(self.closeCalWidget)
            self.childWindow.selectionChanged.connect(self.closeCalWidget)
            self.childWindow.setVisible(True)

    def closeCalWidget(self):
        self.childWindow.setVisible(False)

    def updateDate(self):
        try:
            selectedDate = self.childWindow.selectedDate().toString(dateFormat)
            self.setText(selectedDate)
            self.setFocus()
        except:
            pass

    def dateChanged(self):
        # Auto-add date separators
        global oldDate
        try:
            oldDate
        except NameError:
            oldDate = ''
        newDate = self.text().replace('/', '')
        if len(newDate) > len(oldDate):
            if len(newDate) == 1:
                output_date = newDate
            elif len(newDate) == 2:
                output_date = newDate + '/'
            elif len(newDate) == 3:
                output_date = newDate[:2] + '/' + newDate[2:]
            elif len(newDate) == 4:
                output_date = newDate[:2] + '/' + newDate [2:4] + '/'
            elif len(newDate) > 4:
                output_date = newDate[:2] + '/' + newDate [2:4] + '/' + newDate[4:8]
        else:
            output_date = self.text()
        oldDate = newDate
        self.setText(output_date)

        # if date is valid, update calendar widget to match
        try:
            if len(newDate) > 0:
                self.childWindow.setVisible(True)
        except:
            pass
        try:
            self.childWindow.setSelectedDate(QtCore.QDate.fromString(output_date, dateFormat))
        except:
            pass

class HHHTextEditWidget(QtWidgets.QTextEdit):

    def __init__(self, parent, objName=None, placeholderText=None, readOnly=False, *args, **kwargs):
        super(HHHTextEditWidget, self).__init__(*args, **kwargs)  
        self.parent = parent
        self.setStyleSheet(design_textEdit + design_tableVerticalScrollBar)
        if objName:
            self.setObjectName(objName)
        if placeholderText:
            self.setPlaceholderText(str(placeholderText))
        if readOnly is True:
            self.setReadOnly(True)
        self.setTabChangesFocus(True)
    
    def focusInEvent(self, event):
        super(HHHTextEditWidget, self).focusInEvent(event)
        self.selectAll()
        
    def focusOutEvent(self, event):
        cursor = self.textCursor()
        cursor.movePosition(QtGui.QTextCursor.End)
        self.setTextCursor(cursor)
        super(HHHTextEditWidget, self).focusOutEvent(event)

class HHHComboBox(QtWidgets.QComboBox):
    
    def __init__(self, parent, objName=None, placeholderText=None, readOnly=False, *args, **kwargs):
        super(HHHComboBox, self).__init__(*args, **kwargs)
        self.parent = parent
        self.setStyleSheet(design_comboBox)
        if objName:
            self.setObjectName(objName)
        if placeholderText:
            self.setEditable(True)
            self.setCurrentIndex(-1)
            self.lineEdit().setPlaceholderText(placeholderText)
        if readOnly is True:
            self.lineEdit().setReadOnly(readOnly)
        self.installEventFilter(self)
        for child in self.findChildren(QtWidgets.QListView):
            child.installEventFilter(self) # required for the popup list to be effected by eventFilter
        self.highlighted.connect(self.setCurrentIndex)

    def eventFilter(self, obj, event):
        if isinstance(obj, QtWidgets.QComboBox):
            if event.type() == QtCore.QEvent.KeyPress:
                pass
            return super(HHHComboBox, self).eventFilter(obj, event)
        elif isinstance(obj, QtWidgets.QListView):
            if event.type() == QtCore.QEvent.KeyPress:
                if event.key() in [QtCore.Qt.Key_Return, QtCore.Qt.Key_Enter, QtCore.Qt.Key_Tab, QtCore.Qt.Key_Backtab]:
                    self.hidePopup()
                    event.ignore()
                    return True
            return super().eventFilter(obj, event)

    def focusInEvent(self, event):
        if event.reason() in [QtCore.Qt.TabFocusReason, QtCore.Qt.BacktabFocusReason] and self.count() > 0: 
            setIndex = True if self.currentIndex() < 0 else False            
            self.showPopup()
            if setIndex:
                self.setCurrentIndex(-1)
            self.lineEdit().setFocus(QtCore.Qt.OtherFocusReason)