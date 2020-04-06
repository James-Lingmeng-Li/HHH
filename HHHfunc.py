""" Functions

This is where important functions and methods go, such as sql queries and connections. Most of the 'meat' of the code should in principle go here. If not, consider moving to this module.

The remaining modules should ideally be PyQt5 code only
"""

import pyodbc, pandas, concurrent.futures, time, docx, re, os, shutil, numpy, datetime, subprocess, atexit, threading, copy, traceback, json, psutil, ctypes, win32process
# Some issues with pip install of pywin32. Try pip uninstall then pip install pyqin32==224 and then run: python pywin32_postinstall.py -install 
import win32com.client as win32

from argon2 import PasswordHasher
from cryptography.fernet import Fernet
from dateutil.relativedelta import relativedelta
import xml.etree.ElementTree as et

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

from sqlalchemy import create_engine
from PyQt5 import QtCore, QtWidgets

import HHHconf

backgroundProcesses = [] # store PID strings created by HHH, so exit clean-up can occur

@atexit.register # this decorator registers functions to exit clean-up
def onExit_generic():
    print('Closing...')
    # Close leftover HHH initiated processes
    for process in backgroundProcesses:
        try:
            del p
        except:
            pass
    # # Close any background wordSessions #TODO remove. This killed all word instances including user ones too.
    # for p in psutil.process_iter():
    #     if 'WINWORD' in p.name(): 
    #         p.kill()

    # Delete mailMerge uploaded doc copies
    try:
        shutil.rmtree(HHHconf.temp_dir + '\\_uploaded')
        print('deleted temporary documents uploaded for merge')
    except:
        pass
    print('X_X')

# decorator for AHK macros to disable USER INPUT
def disableInputDuringExecution(func): 
    def wrapper(*args, **kwargs):
        # create thread that blocks input through AHK hotkeys (cannot use keyboard module, as suppressing a key in this module prevents the key event from being sent)
        try:
            thread_inputOff = threading.Thread(target=turnOffInput)
            thread_inputOff.start()
            print('input BLOCKED')
            return func(*args, **kwargs)
        except:
            pass
        finally:
            subprocess.Popen([HHHconf.toggleInput_dir, 'ON'])
            print('input UNBLOCKED')

    def turnOffInput():
        p = subprocess.Popen([HHHconf.toggleInput_dir, 'OFF'], stdout=subprocess.PIPE)
        while True:
            line = p.stdout.readline()
            if not line:
                break
            output = str(line.strip().decode('utf-8'))
            if output == 'User-initiated exit':
                print('pressed ENTER, thread interrupted')
                break
        return
    return wrapper

# add this decorator to get a function's execution time printed
def printExecTime(func):
    def wrapper(*args, **kwargs):
        startTime = datetime.datetime.now()
        func(*args, **kwargs)
        print(str(func) + ' execution time = ' + str(datetime.datetime.now() - startTime))
    return wrapper

def is_number(s):
    try:
        float(s)
        return True
    except ValueError:
        return False

def is_int(s):
    try:
        int(s)
        return True
    except ValueError:
        return False

def is_datetime(x, format):
    import datetime
    try:
        datetime.datetime.strptime(x, format)
        return True
    except ValueError:
        return False

def is_ABN(x):
    """
    FROM ABR WEBSITE 
    To verify an ABN: https://abr.business.gov.au/Help/AbnFormat
    Subtract 1 from the first (left-most) digit of the ABN to give a new 11 digit number
    Multiply each of the digits in this new number by a "weighting factor" based on its position as shown in the table below
    Sum the resulting 11 products
    Divide the sum total by 89, noting the remainder
    If the remainder is zero the number is a valid ABN
    """
    divisibleBy = 89
    list_positionWeight = [10, 1, 3, 5, 7, 9, 11, 13, 15, 17, 19]
    try:
        abn = str(x).replace(' ', '')
        if len(abn) != 11:
            return False
        specialSum = 0
        for i in range(len(abn)):
            if not is_int(abn[i]):
                return False
            if i == 0:
                specialSum  += list_positionWeight[i] * (int(abn[i]) - 1)
            else:
                specialSum  += list_positionWeight[i] * int(abn[i])
        if specialSum % divisibleBy == 0:
            return True
    except:
        pass
    return False

def is_ACN(x):
    """
    FROM ASIC WEBSITE 
    To verify an ACN: https://asic.gov.au/for-business/registering-a-company/steps-to-register-a-company/australian-company-numbers/australian-company-number-digit-check/
    """
    divideBy = 10
    list_positionWeight = [8, 7, 6, 5, 4, 3, 2, 1]
    try:
        acn = str(x).replace(' ', '')
        if len(acn) != 9:
            return False
        specialSum = 0
        checkDigit = int(acn[8]) # extract last ACN digit
        for i in range(8): # weight-sum remaining digits
            if not is_int(acn[i]):
                return False
            specialSum  += list_positionWeight[i] * int(acn[i])
            
        specialSumCompliment = divideBy - (specialSum % divideBy)
        if specialSumCompliment == 10:
            specialSumCompliment = 0
        if specialSumCompliment == checkDigit:
            return True
    except:
        pass
    return False

# Converts a number to equivalent excel lettered column
def numberToXLCol(n):
    string = ''
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string

def getScaleFactor():
    return ctypes.windll.shcore.GetScaleFactorForDevice(0) / 100

class mainEngine():
    @staticmethod
    def verifyPermissions(permissionCheck):
        df_loginDetails = mainEngine.getLoginDetails(HHHconf.username_PC)
        userAuthority = df_loginDetails.at[0, HHHconf.userAdmin_authority]
        if permissionCheck in HHHconf.dict_permissions[userAuthority]:
            return True
        else:
            return False

    # TODO: Check user authorities to allow or deny access to the sql server 
    @staticmethod
    def establishSqlConnection(authority=0):
        return pyodbc.connect('DRIVER={SQL Server};SERVER=' + HHHconf.sqlServer + ';DATABASE=' + HHHconf.sqlDatabase + ';UID=' + HHHconf.sqlUsername + ';PWD=' + HHHconf.sqlPassword)

    @staticmethod
    def oneWayEncryptPattern(string:str): 
        return string + str(len(string))

    @staticmethod
    def oneWayEncrypt(string:str):
        return PasswordHasher().hash(mainEngine.oneWayEncryptPattern(string))

    @staticmethod
    def twoWayEncrypt(string:str):
        return Fernet(open(HHHconf.fernetKey_dir, 'r').read().encode('ascii')).encrypt(string.encode('ascii'))
    
    @staticmethod
    def twoWayDecrypt(string:str):
        return Fernet(open(HHHconf.fernetKey_dir, 'r').read().encode('ascii')).decrypt((string).encode('ascii')).decode('ascii')

    @staticmethod
    def getLoginDetails(username:str=None):
        if username:
            job_getLoginDetails = ('SELECT * FROM ' + HHHconf.db_userAdmin_name + ' WHERE ' + HHHconf.dict_toSQL_userAdmin[HHHconf.userAdmin_username] + ' = \'' + username + '\'')
        else:
            job_getLoginDetails = ('SELECT * FROM ' + HHHconf.db_userAdmin_name)
        with mainEngine.establishSqlConnection() as cnxn:
            df_loginDetails = pandas.read_sql(job_getLoginDetails, cnxn)
        
        df_loginDetails = df_loginDetails.rename(HHHconf.dict_fromSQL_userAdmin, axis='columns')
        return df_loginDetails

    @staticmethod
    def updateLoginDetails(username:str, data=None):
        # If New User DB entry
        if mainEngine.getLoginDetails(username).empty:
            with mainEngine.establishSqlConnection() as cnxn:
                with cnxn.cursor() as cursor:
                    password = mainEngine.oneWayEncrypt(HHHconf.defaultPassword)
                    authority = 'HHH_default'
                    job_addNewUser = 'INSERT INTO ' + HHHconf.db_userAdmin_name + '(' + HHHconf.dict_toSQL_userAdmin[HHHconf.userAdmin_username] + ', ' + HHHconf.dict_toSQL_userAdmin[HHHconf.userAdmin_password] + ', ' + HHHconf.dict_toSQL_userAdmin[HHHconf.userAdmin_authority] + ') VALUES(?, ?, ?)'
                    cursor.execute(job_addNewUser, username, password, authority)
                    cnxn.commit()
        if data is None:
            data = {}
        with mainEngine.establishSqlConnection() as cnxn:
            with cnxn.cursor() as cursor:
                for key, value in data.items():
                    try:
                        job_updateLoginDetails = 'UPDATE ' + HHHconf.db_userAdmin_name + ' SET ' + HHHconf.dict_toSQL_userAdmin[key] + ' = ? WHERE ' + HHHconf.dict_toSQL_userAdmin[HHHconf.userAdmin_username] + ' = ?'
                        cursor.execute(job_updateLoginDetails, value, username)
                        cnxn.commit()
                    except:
                        pass

    @staticmethod
    def getAppCredentials(username:str):
        job_getAppCredentials = 'SELECT * FROM ' + HHHconf.db_userCredentials_name +  ' WHERE ' + HHHconf.dict_toSQL_userCredentials[HHHconf.userCredentials_user] + ' = \'' + username + '\''
        job_getAppCredentialsHalo = 'SELECT * FROM ' + HHHconf.db_userCredentials_name +  ' WHERE ' + HHHconf.dict_toSQL_userCredentials[HHHconf.userCredentials_application] + ' LIKE \'HALO%\'' # All DF users should have access to Halo logins
        with mainEngine.establishSqlConnection() as cnxn:
            df_appCredentialsUser = pandas.read_sql(job_getAppCredentials, cnxn) 
            df_appCredentialsHalo = pandas.read_sql(job_getAppCredentialsHalo, cnxn)
        df_appCredentials = pandas.concat([df_appCredentialsUser, df_appCredentialsHalo.reset_index(drop=True)], axis=0)
        df_appCredentials = df_appCredentials.rename(HHHconf.dict_fromSQL_userCredentials, axis='columns')
        return df_appCredentials

    @staticmethod
    def updateAppCredentials(app:str, data:dict, username:str):
        # Updates existing app credentials or inserts new one if not existing
        with mainEngine.establishSqlConnection() as cnxn:
            with cnxn.cursor() as cursor:
                for key, value in data.items():
                    try:
                        usernameRequired = '' if app in HHHconf.list_commonAppNames else ' AND ' + HHHconf.dict_toSQL_userCredentials[HHHconf.userCredentials_user] + ' = \'' + username + '\''
                        job_updateLoginDetails = ('IF EXISTS(SELECT * FROM ' + HHHconf.db_userCredentials_name + ' WHERE ' + HHHconf.dict_toSQL_userCredentials[HHHconf.userCredentials_application] + ' = ?' + usernameRequired + ') UPDATE ' + HHHconf.db_userCredentials_name + ' SET ' + HHHconf.dict_toSQL_userCredentials[key] + ' = ? WHERE ' + HHHconf.dict_toSQL_userCredentials[HHHconf.userCredentials_application] + ' = ?' + usernameRequired + ' ELSE INSERT INTO ' + HHHconf.db_userCredentials_name + '(' + HHHconf.dict_toSQL_userCredentials[HHHconf.userCredentials_user] + ', ' + HHHconf.dict_toSQL_userCredentials[HHHconf.userCredentials_application] + ', ' + HHHconf.dict_toSQL_userCredentials[key] + ') VALUES(?, ?, ?)')
                        cursor.execute(job_updateLoginDetails, app, value, app, username, app, value)
                        cnxn.commit()
                    except:
                        print(traceback.format_exc())

    # Optional TODO: remove if completely deprecated by updateAppCredentials method
    # @staticmethod            
    # def addAppCredentials(data:dict):
    #     strflist_credentialsColumns = ' , '.join(list(HHHconf.dict_toSQL_userCredentials.values()))
    #     with mainEngine.establishSqlConnection() as cnxn:
    #         with cnxn.cursor() as cursor:
    #             job_updateOrAddAppCredentials = ('IF EXISTS(SELECT * FROM ' + HHHconf.db_userCredentials_name + ' WHERE ' + HHHconf.dict_toSQL_userCredentials[HHHconf.userCredentials_user] + ' = ? AND ' + HHHconf.dict_toSQL_userCredentials[HHHconf.userCredentials_application] + ' = ?) UPDATE ' + HHHconf.db_userCredentials_name + ' SET ' + HHHconf.dict_toSQL_userCredentials[HHHconf.userCredentials_username] + ' = ?, ' + HHHconf.dict_toSQL_userCredentials[HHHconf.userCredentials_password] + ' = ?, ' + HHHconf.dict_toSQL_userCredentials[HHHconf.userCredentials_credential1] + ' = ?, ' + HHHconf.dict_toSQL_userCredentials[HHHconf.userCredentials_credential2] + ' = ? WHERE ' + HHHconf.dict_toSQL_userCredentials[HHHconf.userCredentials_user] + ' = ? AND ' + HHHconf.dict_toSQL_userCredentials[HHHconf.userCredentials_application] + ' = ? ELSE INSERT INTO ' + HHHconf.db_userCredentials_name + ' (' + strflist_credentialsColumns + ') VALUES(' + ', '.join('?'*len(HHHconf.dict_toSQL_userCredentials)) + ')')
    #             cursor.execute(job_updateOrAddAppCredentials, data[HHHconf.userCredentials_user], data[HHHconf.userCredentials_application], data[HHHconf.userCredentials_username], data[HHHconf.userCredentials_password], data[HHHconf.userCredentials_credential1], data[HHHconf.userCredentials_credential2], data[HHHconf.userCredentials_user], data[HHHconf.userCredentials_application], data[HHHconf.userCredentials_user], data[HHHconf.userCredentials_application], data[HHHconf.userCredentials_username], data[HHHconf.userCredentials_password], data[HHHconf.userCredentials_credential1], data[HHHconf.userCredentials_credential2])
    #             cnxn.commit()
    
    @staticmethod
    def deleteAppCredentials(username:str, app:str):
         with mainEngine.establishSqlConnection() as cnxn:
            with cnxn.cursor() as cursor:
                job_deleteAppCredentials = 'DELETE FROM ' + HHHconf.db_userCredentials_name + ' WHERE ' + HHHconf.dict_toSQL_userCredentials[HHHconf.userCredentials_user] + ' = ? AND ' + HHHconf.dict_toSQL_userCredentials[HHHconf.userCredentials_application] + ' = ?'
                cursor.execute(job_deleteAppCredentials, username, app)
                cnxn.commit()

    @staticmethod
    def requestPassword(parent, user:str, title:str='Security', dialogText:str='Please enter your password:'):
        attempts = 0
        while True:
            incorrectPasswordText = '' if attempts < 1 else '\nIncorrect Password. Please try again.'
            attemptsText = '' if attempts < 5 else '\n\n...It\'s time to request a password reset'
            enteredPassword, confirmed = QtWidgets.QInputDialog.getText(parent, title, dialogText + incorrectPasswordText + attemptsText, QtWidgets.QLineEdit.Password, '')
            if confirmed is False:
                return False
            if mainEngine.validatePassword(user, enteredPassword) is False:
                attempts  += 1
                continue
            else:
                return True

    @staticmethod
    def validatePassword(username:str, passwordString:str):
        try:
            df_loginDetails = mainEngine.getLoginDetails(username)
            currentPasswordHashed = df_loginDetails.at[0, HHHconf.userAdmin_password]
            PasswordHasher().verify(currentPasswordHashed, mainEngine.oneWayEncryptPattern(passwordString))
            return True
        except:
            return False
    
    @staticmethod
    def sendInternalNotification():
        # send auto email to user, as well as add the notfication to their notifications (TODO: create notification center) to view when logged in
        pass

    @staticmethod
    def newChromeSession(defaultDir=HHHconf.downloads_dir, printOrSave='save', detach=False):
        destination = 'Save as PDF' if printOrSave == 'save' else HHHconf.printerID
        appState = {
                'recentDestinations': [{
                        'id': destination, 
                        'origin': 'local', 
                        'account':''
                }], 
                'selectedDestinationId': destination, 
                'version': 2
        }
        chrome_options = webdriver.ChromeOptions()
        profile = {'printing.print_preview_sticky_settings.appState': json.dumps(appState), 'download.default_directory': defaultDir + '\\\\', 'download.directory_upgrade':True, 'download.prompt_for_download':False, 'profile.default_content_setting_values.automatic_downloads': 2, 'safebrowsing.enabled':True}
        chrome_options.add_experimental_option('prefs', profile)
        if detach is True: # Keeps browser open after code completes execution
            chrome_options.add_experimental_option('detach', True)
        chrome_options.add_argument('--kiosk-printing') # This was for auto-print, but the save as pdf recentDestinations is not working at the moment 
        chrome_options.add_argument('--start-maximized')
        chrome_options.add_argument('--disable-popup-blocking')
        driver = webdriver.Chrome(HHHconf.chromeDriver_dir, chrome_options=chrome_options)
        return driver

    @staticmethod
    def addEvent(dict_event):
        with mainEngine.establishSqlConnection() as cnxn:
            cursor = cnxn.cursor()
            # add timestamp, as some functions do not do that
            if dict_event[HHHconf.events_EventTime] is None:
                dict_event[HHHconf.events_EventTime] = datetime.datetime.now()
            columnNames = ', '.join([HHHconf.dict_toSQL_Events[x] for x in list(dict_event.keys())])
            placeholders = ', '.join(['?'] * len(dict_event))
            job_addEvent = 'INSERT INTO ' + HHHconf.db_Events_name + ' (' + columnNames + ') OUTPUT INSERTED.EventID VALUES (' + placeholders + ')'
            cursor.execute(job_addEvent, list(dict_event.values()))
            newEventID = cursor.fetchone()[0]
            cnxn.commit()
        return str(newEventID)

class mailMergeEngine():

    @staticmethod
    def newOutlookSession():
        outlook = win32.DispatchEx('Outlook.Application')
        backgroundProcesses.append(outlook)
        return outlook
        
    @staticmethod
    def newWordSession():
        word = win32.DispatchEx('Word.Application')
        backgroundProcesses.append(word)
        return word

    @staticmethod
    def PrimarySmtpAddress():
        outlook = mailMergeEngine.newOutlookSession()
        emailAddress = str(outlook.Session.CurrentUser.AddressEntry.GetExchangeUser().PrimarySmtpAddress)  
        del outlook
        return emailAddress

    @staticmethod
    def get_valid_filename(s:str):
        """
        Return the given string converted to a string that can be used for a clean
        filename without the extension. Remove leading and trailing spaces; convert other spaces to
        underscores; and remove anything that is not an alphanumeric, dash, 
        underscore, and removes dots too.
        """       
        s = str(s).strip().replace('.', '')
        s = s.replace(' ', '_')
        s = s.replace('__','_')
        return re.sub(r'(?u)[^-\w.]', '', s)
        
    @staticmethod
    def start(threadCompleteEvent, mergeData, mergeOptions, mergeDocuments=None, varToCols=None, username=None):
        startTime = time.time()                
        # Connect to Outlook
        try:
            outlook = mailMergeEngine.newOutlookSession()
        except:
            print('Error: Failed to Open Outlook')
            return

        # Connect to Word
        try:
            word = mailMergeEngine.newWordSession()
        except:
            print('Error: Failed to hook Microsoft Word')
            return

        mergeData['dataframe'].fillna(value='', inplace=True)
        dict_mergeData = mergeData['dataframe'].to_dict('index')
        list_mergeDataTuples = dict_mergeData.items()

        # Create threads
        with concurrent.futures.ThreadPoolExecutor(max_workers=10) as executor:
            results = executor.map(lambda x: mailMergeEngine.mergeThread(x, mergeOptions=mergeOptions, mergeDocuments=mergeDocuments, varToCols=varToCols, outlookSession=outlook, wordSession=word, username=username), list_mergeDataTuples)
            
        # Delete temp files not elected for saving 
        try:
            shutil.rmtree(HHHconf.temp_dir + '\\_merged')
        except:
            pass
        finally:
            print(f'done in {startTime - time.time()} seconds: {list(results)}')
            try:
                word.Quit()
            except:
                print('could not close Word')
                print(traceback.format_exc())
            finally:
                del word
                del outlook
            # TODO: Save merge results somewhere

        # trigger completionThread event to be picked up by thread in main method, by setting the event to be True
        threadCompleteEvent.set() 

    @staticmethod
    def mergeThread(mergeInstance, mergeOptions=None, mergeDocuments=None, varToCols=None, outlookSession=None, wordSession=None, username=None):
        list_attachFiles = [] # used to store file paths to attach to email
        # For each df row, save files, compose email, attach files to email and send
        uniqueID = mergeInstance[0]
        dict_mergeInstance = mergeInstance[1]

        # merge documents with data then save as doc or pdf
        for doc in mergeDocuments.values():
            docName = os.path.basename(doc['path'])
            docExt = os.path.splitext(docName)[1]
                    
            # Create instance copy first, to use for merging
            xStr = lambda s: '' if s is None else mailMergeEngine.get_valid_filename(s)
            try:
                subDir = xStr(dict_mergeInstance[mergeOptions['save']['subDir']]) + '\\'
            except:
                subDir = ''

            # set save path to temp directory, if save option is set to False
            if doc['mergeOptions']['saveFile'] is True:
                savePath = mergeOptions['save']['dir'] + '\\' + subDir
            else:
                savePath = HHHconf.temp_dir + '\\_merged\\' + subDir
            os.makedirs(savePath, exist_ok=True)
            try:
                suffix = xStr(dict_mergeInstance[doc['mergeOptions']['uniqueSuffix']])
            except:
                suffix = ''

            docPath = savePath + xStr(doc['mergeOptions']['saveAsName']) + '_' + suffix + '_' + str(uniqueID) + docExt
            docCopied = str(shutil.copy2(doc['path'], docPath))

            try:
                uploadedDocument = docx.Document(docCopied)
            except:
                pass
            list_varsReplaced = []
            if uploadedDocument:
                try:
                    # If doc has vars, merge with data before saving
                    if len(doc['docVars']) != 0:
                        list_varsReplaced = mailMergeEngine.replaceVarsInDocObj(uploadedDocument, varToCols, dict_mergeInstance, runningList=list_varsReplaced)
                        for sctn in uploadedDocument.sections:
                            list_varsReplaced = mailMergeEngine.replaceVarsInDocObj(sctn.header, varToCols, dict_mergeInstance, runningList=list_varsReplaced)
                            list_varsReplaced = mailMergeEngine.replaceVarsInDocObj(sctn.footer, varToCols, dict_mergeInstance, runningList=list_varsReplaced)
                            if sctn.different_first_page_header_footer:
                                list_varsReplaced = mailMergeEngine.replaceVarsInDocObj(sctn.first_page_header, varToCols, dict_mergeInstance, runningList=list_varsReplaced)
                                list_varsReplaced = mailMergeEngine.replaceVarsInDocObj(sctn.first_page_footer, varToCols, dict_mergeInstance, runningList=list_varsReplaced)
                            if uploadedDocument.settings.odd_and_even_pages_header_footer:
                                # Odd headers/footers captured by .header/.footer objects by default 
                                list_varsReplaced = mailMergeEngine.replaceVarsInDocObj(sctn.even_page_header, varToCols, dict_mergeInstance, runningList=list_varsReplaced)
                                list_varsReplaced = mailMergeEngine.replaceVarsInDocObj(sctn.even_page_footer, varToCols, dict_mergeInstance, runningList=list_varsReplaced)
                        uploadedDocument.save(docCopied)

                        # Save as PDF if specified
                        if doc['mergeOptions']['saveAsExtension'] == '.pdf':
                            docPath = savePath + xStr(doc['mergeOptions']['saveAsName']) + '_' + suffix + '_' + str(uniqueID) + doc['mergeOptions']['saveAsExtension']
                            mailMergeEngine.docToPDF(docCopied, docPath, wordSession)
                except:
                    print('ERROR ENCOUNTERED IN DOC CONVERSION: -- ' + str(traceback.format_exc()))
            # Simply save file if not a word doc
            else:
                print(traceback.format_exc())
                try:
                    suffix = xStr(dict_mergeInstance[doc['mergeOptions']['uniqueSuffix']])
                except:
                    suffix = ''
                docPath = savePath + xStr(doc['mergeOptions']['saveAsName']) + '_' + suffix + '_' + str(uniqueID) + doc['mergeOptions']['saveAsExtension']
                shutil.copy2(doc['path'], docPath)
            if doc['mergeOptions']['attachToEmail'] is True:
                list_attachFiles.append(docPath)

        # email
        dict_recipients = {
            'to': None, 
            'cc': None, 
            'bcc': None
        }
        try:
            dict_recipients['to'] = dict_mergeInstance[mergeOptions['send']['to']]
        except:
            # End thread if email not being sent
            return (str(uniqueID), None, True)
        try:
            dict_recipients['cc'] = dict_mergeInstance[mergeOptions['send']['cc']]
        except:
            pass
        try:
            dict_recipients['bcc'] = dict_mergeInstance[mergeOptions['send']['bcc']]
        except:
            pass

        #  replace vars in subject and body
        try:
            sub = mergeOptions['send']['subject']
        except:
            sub = ''
        varsInSubject = list(set(re.findall('<<(.+?)>>', sub)))
        for var in varsInSubject:
            try:
                stringToInsert = str(dict_mergeInstance[var]).strip() # Remove surrounding whitespace
            
            # Insert default value if string is empty (i.e. original excel cell had no real data)
                if len(stringToInsert) == 0 and varToCols[var]:
                    stringToInsert = str(varToCols[var]['defaultValue']).strip()  
                sub = sub.replace('<<' + var + '>>', stringToInsert)
            except:
                continue
        try:
            bod = mergeOptions['send']['body']
        except:
            bod = ''
        varsInBody = list(set(re.findall('&lt;&lt;(.+?)&gt;&gt;', bod))) # html for '<<' and '>>'
        for var in varsInBody:
            try:
                stringToInsert = str(dict_mergeInstance[var]).strip() # Remove surrounding whitespace
                # Insert default value if string is empty (i.e. original excel cell had no real data)
                if len(stringToInsert) == 0 and varToCols[var]:
                    stringToInsert = str(varToCols[var]['defaultValue']).strip()
                bod = bod.replace('&lt;&lt;' + var + '&gt;&gt;', stringToInsert)
            except:
                continue
        emailSentTo = mailMergeEngine.sendEmail(outlookSession=outlookSession, sender=mailMergeEngine.PrimarySmtpAddress(), recipient=dict_recipients, subject=sub, body=bod, attachments=list_attachFiles)

        # print any files selected to print - Unfinished

        # Return results of merge instance
        return (str(uniqueID), emailSentTo)

    @staticmethod
    def replaceVarsInDocObj(docObj, varToCols, mergeInstance, runningList=None):
        if runningList is None:
            runningList = []
        # captures tables in tables 
        for tbl in docObj.tables:
            for row in tbl.rows:
                for cell in row.cells: 
                    runningList = mailMergeEngine.replaceVarsInDocObj(cell, varToCols, mergeInstance, runningList=runningList)

        for pgh in docObj.paragraphs:
            # Skip if pgh text does not contain variable
            checkIfAnyVars = list(set(re.findall('<<(.+?)>>', pgh.text)))
            if len(checkIfAnyVars) == 0:
                continue
            
            # for each run, check if any full variables are there. If so, replace with data string. Then check if partial variables are in the run, if so, check subsequent runs to see when the variable is complete, remove the variable components from these runs and change the first run to include the variable data at the end of the run
            runs = pgh.runs
            for i in range(len(runs)):
                # Check for full var matches in this run
                foundVars = list(set(re.findall('<<(.+?)>>', runs[i].text)))
                for var in foundVars:
                    if varToCols[var]['excludeFromMerge'] is False:
                        stringToInsert = str(mergeInstance[varToCols[var]['assignColumn']]).strip() # Remove surrounding whitespace
                        # Insert default value if string is empty (i.e. original excel cell had no real data)
                        if len(stringToInsert) == 0:
                            stringToInsert = str(varToCols[var]['defaultValue']).strip()
                        runs[i].text = runs[i].text.replace('<<' + var + '>>', stringToInsert)
                runningList = runningList + foundVars
               
                # Check for partial vars that start in this run
                varAcrossRunsFound = False                
                potentialPos = runs[i].text.find('<')
                # if '<' is found in the run
                if potentialPos >= 0:
                    startOfVar = (i, potentialPos)
                    # and if next character is also '<'
                    try:
                        nextChar = runs[i].text[potentialPos + 1]
                    except:
                        try:
                            nextChar = runs[i + 1].text[0] # error means, code is on the last run in the paragraph
                        except:
                            break      
                    if nextChar == '<':
                        # then try to find one trailing '>' in remaining runs
                        for j in range(i, len(runs)):
                            if j == i:
                                potentialTrailingPos = runs[j].text.find('>', potentialPos)
                            else:
                                potentialTrailingPos = runs[j].text.find('>')
                            # if '>' is found in the run
                            if potentialTrailingPos >= 0:
                                # and if next character is also '>'
                                try:
                                    nextTrailingChar = runs[j].text[potentialTrailingPos + 1]
                                    endOfVar = (j, potentialTrailingPos + 1)
                                except:
                                    try:
                                        nextTrailingChar = runs[j + 1].text[0] # error means, code is on the last run in the paragraph
                                        endOfVar = (j + 1, 0)
                                    except:
                                        break    
                                if nextTrailingChar == '>':
                                    # get variable found spanning across runs
                                    varAcrossRunsFound = runs[startOfVar[0]].text[startOfVar[1]:]
                                    try:
                                        for k in range(startOfVar[0] + 1, endOfVar[0]):
                                            varAcrossRunsFound = varAcrossRunsFound + runs[k].text
                                    except:
                                        pass
                                    varAcrossRunsFound = varAcrossRunsFound + runs[endOfVar[0]].text[:endOfVar[1] + 1]
                                    varAcrossRunsFound = varAcrossRunsFound.replace('<', '')
                                    varAcrossRunsFound = varAcrossRunsFound.replace('>', '')
                                    # If var found spanning across runs, replace this found var with appropriate data
                                    try:
                                        if varToCols[varAcrossRunsFound]['excludeFromMerge'] is False:        
                                            stringToInsert = str(mergeInstance[varToCols[varAcrossRunsFound]['assignColumn']]).strip() # Remove surrounding whitespace
                                            # Insert default value if string is empty (i.e. original excel cell had no real data)
                                            if len(stringToInsert) == 0:
                                                stringToInsert = str(varToCols[varAcrossRunsFound]['defaultValue']).strip()
                                            runs[startOfVar[0]].text = runs[startOfVar[0]].text[:startOfVar[1]] + stringToInsert # append data to end of first run
                                            try:
                                                for k in range(startOfVar[0] + 1, endOfVar[0]):
                                                    runs[k].clear() # remove any middle-laying runs
                                            except:
                                                pass
                                            runs[endOfVar[0]].text = runs[endOfVar[0]].text[endOfVar[1] + 1:]
                                            runningList.append(varAcrossRunsFound)
                                    except KeyError:
                                        pass
        runningList = list(set(runningList))
        return runningList 

    @staticmethod
    def docToPDF(docPath, pdfPath, wordSession):
        doc = wordSession.Documents.Open(docPath, ReadOnly=True)
        while True:
            try:
                doc.SaveAs(pdfPath, FileFormat=17)
            except:
                print('waiting to save: ' + str(doc))
            else:
                break
        while True:
            try:
                doc.Close()
            except:
                print('waiting to close: ' + str(doc))
            else:
                break
        return True

    @staticmethod
    def sendEmail(outlookSession, sender, recipient={'to':None, 'cc':None, 'bcc':None}, subject='', body='', attachments=[], displayOnly=False):
        if recipient['to'] is None:
            return None
        try:
            mail = outlookSession.CreateItem(0)
            mail.SentOnBehalfOfName = str(sender)
            mail.To = recipient['to']
            if recipient['cc'] is not None:
                mail.cc = recipient['cc']
            if recipient['bcc'] is not None:
                mail.bcc = recipient['bcc']
            mail.Subject =  subject
            mail.GetInspector # Creates signature in html body
            
            signatureBody = mail.HTMLBody
            pos = signatureBody.find('>', signatureBody.find('<body'))
            mail.HTMLBody = signatureBody[:pos + 1] + body + signatureBody[pos + 1:]
            for attachment in attachments:
                mail.Attachments.Add(attachment)
            if displayOnly is True:
                return mail.Display(False) # TRUE makes the code hang until the mail window is closed, FALSE does not interrupt code, and still displays the window
            else:
                return mail.send
        except Exception as e:
            return print(traceback.format_exc())

class debtorFinanceEngine():
    @staticmethod
    def getClientDetails(selectedClients:list=None, includeAQ=True, includeB2B=True):
        job_getDFClients = ('''
        SELECT * FROM ''' + HHHconf.df_DFClientInfo_name + '''
        ''')
        df_debtorClients = pandas.read_sql(job_getDFClients, mainEngine.establishSqlConnection())
        df_debtorClients.rename(columns=HHHconf.dict_fromSQL_DFClientInfo, inplace=True)
        if includeAQ is False:
            df_debtorClients = df_debtorClients[df_debtorClients[HHHconf.DFClientInfo_AQClientNumber].str[:1] != '0']
        if includeB2B is False:
            df_debtorClients = df_debtorClients[df_debtorClients[HHHconf.DFClientInfo_AQClientNumber].str[:1] != '9']
        if isinstance(selectedClients, list):
            df_debtorClients = df_debtorClients[df_debtorClients[HHHconf.DFClientInfo_AQClientNumber].isin(selectedClients)]
        return df_debtorClients

    @staticmethod
    def getDebtorReceipts(date=None):
        job_getDebtorReceipts = ('''
        SELECT 
            b.[''' + HHHconf.dict_toSQL_DFClientInfo[HHHconf.DFClientInfo_AQClientNumber] + '''], 
            a.[''' + HHHconf.dict_toSQL_debtorReceipts[HHHconf.debtorReceipts_client] + '''], 
            a.[''' + HHHconf.dict_toSQL_debtorReceipts[HHHconf.debtorReceipts_value] + '''], 
            a.[''' + HHHconf.dict_toSQL_debtorReceipts[HHHconf.debtorReceipts_status] + ''']
        FROM [mPower].[dbo].[''' + HHHconf.db_debtorReceipts_name + '''] a INNER JOIN [mPower].[dbo].[''' + HHHconf.df_DFClientInfo_name + '''] b on a.[''' + HHHconf.dict_toSQL_debtorReceipts[HHHconf.debtorReceipts_client] + '''] = b.[''' + HHHconf.dict_toSQL_DFClientInfo[HHHconf.DFClientInfo_AQClientName] + ''']
        WHERE a.[''' + HHHconf.dict_toSQL_debtorReceipts[HHHconf.debtorReceipts_date] + '''] = \'''' + date.strftime(HHHconf.dateFormat3) + '''\' AND a.[''' + HHHconf.dict_toSQL_debtorReceipts[HHHconf.debtorReceipts_value] + '''] IS NOT NULL
        ORDER BY b.[''' + HHHconf.dict_toSQL_DFClientInfo[HHHconf.DFClientInfo_AQClientNumber] + ''']
        ''')
        df_debtorReceipts = pandas.read_sql(job_getDebtorReceipts, mainEngine.establishSqlConnection())
        df_debtorReceipts.rename(columns=HHHconf.dict_fromSQL_debtorReceipts, inplace=True)
        df_debtorReceipts.rename(columns=HHHconf.dict_fromSQL_DFClientInfo, inplace=True)
        return df_debtorReceipts

    @staticmethod
    def updateDebtorReceiptStatus(oldStatus, newStatus, date, clientNumber=None):
        
        oldStatus = ' is NULL' if oldStatus is None else ' = ' + '\'' + oldStatus + '\''
        newStatus = 'NULL' if newStatus is None else '\'' + newStatus + '\''
        clientNumber = 'b.' + HHHconf.dict_toSQL_DFClientInfo[HHHconf.DFClientInfo_AQClientNumber] if clientNumber is None else '\'' + clientNumber + '\''
        job_updateDebtorReceiptStatus = ('''
        UPDATE
            ''' + HHHconf.db_debtorReceipts_name + '''
        SET
            ''' + HHHconf.dict_toSQL_debtorReceipts[HHHconf.debtorReceipts_status] + ''' = ''' + newStatus + '''
        FROM
            ''' + HHHconf.db_debtorReceipts_name + ''' a 
            INNER JOIN ''' + HHHconf.df_DFClientInfo_name + ''' b ON a.[''' + HHHconf.dict_toSQL_debtorReceipts[HHHconf.debtorReceipts_client] + ''']=b.[''' + HHHconf.dict_toSQL_DFClientInfo[HHHconf.DFClientInfo_AQClientName] + ''']
        WHERE
            a.''' + HHHconf.dict_toSQL_debtorReceipts[HHHconf.debtorReceipts_date] + ''' = ? AND
            a.''' + HHHconf.dict_toSQL_debtorReceipts[HHHconf.debtorReceipts_status] + oldStatus + ''' AND
            b.''' + HHHconf.dict_toSQL_DFClientInfo[HHHconf.DFClientInfo_AQClientNumber] + ''' = ''' + clientNumber + '''
        ''')
        date = date.strftime(HHHconf.dateFormat3)
        with mainEngine.establishSqlConnection() as cnxn:
            with cnxn.cursor() as cursor:
                cursor.execute(job_updateDebtorReceiptStatus, date)
                cnxn.commit()

    class segmentsWorker(QtCore.QObject):
        signal_finished = QtCore.pyqtSignal(object) # pass a HHHconf.dict_toSQL_Events to write to the audit trail database
        signal_started = QtCore.pyqtSignal()
        signal_passwordUpdated = QtCore.pyqtSignal(str)
        def __init__(self, parent, username, WPCusername, WPCpassword):
            super().__init__()
            self.parent = parent
            self.username = username
            f = open(HHHconf.fernetKey_dir, 'r')
            self.key = f.read().encode('ascii')
            self.WPCusername = WPCusername
            self.WPCpassword = WPCpassword 

        def beginSegments(self):
            # TODO test the password change code (only viable once a month when WPC asks for a password change)
            messageToEmit = dict.fromkeys(HHHconf.dict_toSQL_Events)
            messageToEmit[HHHconf.events_UserName] = str(HHHconf.username_PC)            
            messageToEmit[HHHconf.events_Event] = HHHconf.dict_eventCategories['segments']
            messageToEmit[HHHconf.events_EventTime] = datetime.datetime.now()
            eventDescription = 'Worker Initiated | '
            messageToEmit[HHHconf.events_EventDescription] = eventDescription
            try:
                segmentDate = datetime.date.today().strftime(HHHconf.dateFormat4)
                saveFolderDir = HHHconf.saveSegments_dir + '\\' + segmentDate
                os.makedirs(saveFolderDir, exist_ok=True)

                # Delete old SgReports if they exist
                localDownloadDir = 'C:\\Users\\' + HHHconf.username_PC + '\\Downloads'
                for fileName in os.listdir(localDownloadDir):
                    if str(fileName) == 'SGReport.pdf':
                        os.remove(os.path.join(localDownloadDir, fileName))
                        break
                driver = mainEngine.newChromeSession(defaultDir=HHHconf.downloads_dir, printOrSave='save', detach=False)
                driver.get(HHHconf.link_wpc)
                elem_username = driver.find_element_by_id('CUS')
                elem_username.send_keys(self.WPCusername)
                elem_password = driver.find_element_by_id('PWD')
                elem_password.send_keys(self.WPCpassword)
                button_signIn = driver.find_element_by_id('submit-signin')
                oldURL = driver.current_url     
                button_signIn.click()
                WebDriverWait(driver, 20).until(lambda driver: oldURL != driver.current_url and driver.execute_script('return document.readyState == "complete"'))
                handle_main = driver.window_handles[0]
                # Change password if required
                try:
                    elem_oldPassword = driver.find_element_by_id('OldPassword')
                    elem_oldPassword.send_keys(self.WPCpassword)
                    newWPCpassword = HHHconf.defaultPassword  + datetime.date.today().strftime(HHHconf.dateFormat5)
                    print('new password = ' + newWPCpassword)
                    elem_newPassword = driver.find_element_by_id('NewPassword')
                    elem_newPassword.send_keys(newWPCpassword)
                    elem_confirmPassword = driver.find_element_by_id('ConfirmPassword')
                    elem_confirmPassword.send_keys(newWPCpassword)
                    button_submit = driver.find_element_by_id('btnSubmit')
                    button_submit.click()
                    button_submit = driver.find_element_by_id('btnSubmit')
                    button_submit.click()
                    # Update password stored in DB for WPC app credentials
                    self.signal_passwordUpdated.emit(newWPCpassword)
                    eventDescription += 'Please note your WPC password was changed in the process - please see your settings for the new password |'
                except:
                    pass

                # Navigate from Home to Segments screen
                elem_accounts = driver.find_element_by_link_text('Accounts')
                elem_accounts.click()
                elem_segments = driver.find_element_by_link_text('Segments')
                elem_segments.click()

                # for select clients and date, download their segment report
                for name, code in HHHconf.dict_WPCClients.items():
                    elem_segmentReport = driver.find_element_by_link_text(code['segmentLink'])
                    elem_segmentReport.click()
                    elem_printPreview = driver.find_element_by_name('btnPrint')
                    elem_printPreview.click()

                    # PrintPreview window
                    handle_printPreview = driver.window_handles[1]
                    driver.switch_to.window(handle_printPreview)
                    elem_printPage = driver.find_element_by_name('btnPrinttop')
                    elem_printPage.click()

                    # Find downloaded file and move to appropriate folder + rename
                    for fileName in os.listdir(localDownloadDir):
                        if str(fileName) == 'SGReport.pdf':
                            segmentFile = fileName[:fileName.find('.pdf')] + ' - ' + name + '.pdf'
                            segmentPath = os.path.join(saveFolderDir, segmentFile)
                            shutil.move(os.path.join(localDownloadDir, fileName), segmentPath)
                            break
                    elem_closePrintPreview = driver.find_element_by_name('btnClosetop')
                    elem_closePrintPreview.click()
                    driver.switch_to.window(handle_main)
                    elem_segments = driver.find_element_by_link_text('Segments')
                    elem_segments.click()

                    # send signal to initiate email
                    emailBody = ('''
                    <div>
                        Hi ''' + code['contactName'] + ''', <br><br>Please find your segment account report attached for ''' + segmentDate + '''.<br><br>Kind Regards, 
                    </div>
                    ''')
                    mailMergeEngine.sendEmail(outlookSession=mailMergeEngine.newOutlookSession(), sender=mailMergeEngine.PrimarySmtpAddress(), recipient=code['recipient'], subject='Daily Segment Report', body=emailBody, attachments=[segmentPath], displayOnly=True)
                    eventDescription += 'Segment Report Generated: ' + str(name) + ' | '
            except:
                print(traceback.format_exc())
                eventDescription += 'ERROR: An issue occured whilst downloading Segment Reports or the server timed out (>20s) | '
            finally:
                eventDescription += 'Worker Complete |'
                messageToEmit[HHHconf.events_EventDescription] = eventDescription
                try:
                    driver.quit()
                except:
                    pass
                self.signal_finished.emit(messageToEmit)

    class receiptsWorker(QtCore.QObject):
        # TODO: continue testing for robustness of the code. Once ready, the code needs to be migrated to AQ LIVE, rather than AQ Dev, and tested to see if it still works. Check the HHHconf.aq_dir variable, as well as the actual AHK file (HHHconf.importDebtorReceipts_dir)
        signal_finished = QtCore.pyqtSignal(object) # pass a HHHconf.dict_toSQL_Events to write to the audit trail database
        signal_started = QtCore.pyqtSignal()
        def __init__(self, df_receipts, currentStatus, action, parent, username, password):
            super().__init__()
            self.df_receipts = df_receipts
            self.currentStatus = currentStatus
            self.action = action
            self.parent = parent
            self.username = username
            self.password = password

        @disableInputDuringExecution
        def beginReceipts(self):  
            messageToEmit = dict.fromkeys(HHHconf.dict_toSQL_Events)
            messageToEmit[HHHconf.events_UserName] = str(HHHconf.username_PC)            
            messageToEmit[HHHconf.events_Event] = HHHconf.dict_eventCategories['debtorReceipts']
            messageToEmit[HHHconf.events_EventTime] = datetime.datetime.now()
            eventDescription = 'Worker Initiated | ' + str(self.action) + ' | '
            messageToEmit[HHHconf.events_EventDescription] = eventDescription
            try:
                # Make App clickthrough 
                self.signal_started.emit()

                actionCode = HHHconf.importReceiptsAQCode if self.action == 'import' else HHHconf.reverseReceiptsAQCode
                # Loop dataframe - for each valid debtor receipt, update database with username, then read if username was indeed updated. If so, complete debtor receipt import for that row. This is so multiple users can inititate this function at the same time to split work load.
                df_pendingReceiptsTable = None
                for receiptRow in self.df_receipts.itertuples():
                    # Try to update status to username to prevent others from importing same receipt 
                    debtorFinanceEngine.updateDebtorReceiptStatus(oldStatus=self.currentStatus, newStatus=self.parent.dict_receiptsStatus['pending'], date=self.parent.selectedDate, clientNumber=getattr(receiptRow, HHHconf.DFClientInfo_AQClientNumber))
                    
                    # Read back the commited value to see if user has sole right to process receipt in AQ
                    df_pendingReceiptsTable = debtorFinanceEngine.getDebtorReceipts(date=self.parent.selectedDate) 
                    status = df_pendingReceiptsTable.loc[df_pendingReceiptsTable[HHHconf.DFClientInfo_AQClientNumber] == getattr(receiptRow, HHHconf.DFClientInfo_AQClientNumber), HHHconf.debtorReceipts_status].iloc[0]
                    if status != self.parent.dict_receiptsStatus['pending']:
                        print('Overriden by other user: ' + str(getattr(receiptRow, HHHconf.DFClientInfo_AQClientNumber)))
                        continue
                    
                    # Process receipt (import/reversal) into AQ
                    eventDescription += 'started: ' + str(getattr(receiptRow, HHHconf.DFClientInfo_AQClientNumber)) + ' ' + str(getattr(receiptRow, HHHconf.debtorReceipts_value)) + ' | '
                    p = subprocess.Popen([HHHconf.importDebtorReceipt_dir, self.username, self.password, getattr(receiptRow, HHHconf.DFClientInfo_AQClientNumber), str(getattr(receiptRow, HHHconf.debtorReceipts_value)), self.parent.selectedDateString, actionCode], stdout=subprocess.PIPE)

                    # Read output from ahk, and end loop if user initiates exit
                    importResults = ''
                    while True:
                        line = p.stdout.readline()
                        if not line:
                            break
                        importResults = str(line.strip().decode('utf-8'))
                    if importResults == 'User-initiated exit':
                        print(importResults)
                        raise InterruptedError
                    
                    # Update status from pending to completed status
                    completeStatus = self.parent.dict_receiptsStatus['unallocated'] if self.currentStatus == self.parent.dict_receiptsStatus['allocated'] else (self.parent.dict_receiptsStatus['allocated'] if self.currentStatus == self.parent.dict_receiptsStatus['unallocated'] else self.parent.dict_receiptsStatus['pending'])
                    debtorFinanceEngine.updateDebtorReceiptStatus(oldStatus=self.parent.dict_receiptsStatus['pending'], newStatus=completeStatus, date=self.parent.selectedDate, clientNumber=getattr(receiptRow, HHHconf.DFClientInfo_AQClientNumber))
                    eventDescription += 'completed: ' + str(getattr(receiptRow, HHHconf.DFClientInfo_AQClientNumber)) + ' ' + str(getattr(receiptRow, HHHconf.debtorReceipts_value)) + ' | '
            except InterruptedError:
                eventDescription += 'User-initiated exit | '
            finally:
                eventDescription += 'executionTime: ' + str(datetime.datetime.now() - messageToEmit[HHHconf.events_EventTime])
                messageToEmit[HHHconf.events_EventDescription] = eventDescription
                self.signal_finished.emit(messageToEmit)

    class haloWorker(QtCore.QObject):
        signal_finished = QtCore.pyqtSignal(object)
        signal_started = QtCore.pyqtSignal()
        signal_passwordUpdated = QtCore.pyqtSignal(str)
        def __init__(self, parent, app:str, username:str, password:str):
            super().__init__()
            self.parent = parent
            self.app = app
            self.username = username
            self.password = password

        def haloLogin(self):
            messageToEmit = dict.fromkeys(HHHconf.dict_toSQL_Events)
            messageToEmit[HHHconf.events_UserName] = str(HHHconf.username_PC)            
            messageToEmit[HHHconf.events_Event] = HHHconf.dict_eventCategories['halo']
            messageToEmit[HHHconf.events_EventTime] = datetime.datetime.now()
            eventDescription = 'Worker Initiated | ' + self.app + ' | '
            messageToEmit[HHHconf.events_EventDescription] = eventDescription 
            try:
                driver = mainEngine.newChromeSession(defaultDir=HHHconf.downloads_dir, printOrSave='save', detach=True)
                haloLoginLink = HHHconf.link_halo if self.app == 'HALO' else HHHconf.link_haloDemo
                driver.get(haloLoginLink)
                while True:
                    elem_username = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'usernameField')))
                    elem_username.send_keys(self.username)
                    elem_password = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'passwordField')))        
                    elem_password.send_keys(self.password)
                    button_signIn = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'loginButton')))
                    button_signIn.click()
                    element = WebDriverWait(driver, 20).until(lambda x: x.find_elements_by_id('oldPasswordField') or x.find_elements_by_id('responsiveMainMenuBean-button-1'))[0]
                    if element.get_attribute('id') == 'oldPasswordField':
                        try:
                            elem_oldPassword = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'oldPasswordField')))   
                            elem_oldPassword.send_keys(self.password)
                            newPassword = (str(HHHconf.haloPassword) + datetime.datetime.now().strftime(HHHconf.dateFormat5)) if self.app == 'HALO' else (str(HHHconf.haloDemoPassword) + datetime.datetime.now().strftime(HHHconf.dateFormat5))
                            elem_newPassword = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'newPasswordField1')))
                            elem_newPassword.send_keys(newPassword)
                            elem_confirmPassword = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'newPasswordField2')))
                            elem_confirmPassword.send_keys(newPassword)
                            button_changePassword = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'loginButton')))
                            button_changePassword.click()
                            button_return = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'haloUI-Label')))
                            button_return.click()
                            self.signal_passwordUpdated.emit(newPassword)
                            self.password = newPassword
                            eventDescription += 'Password changed. Check your settings | '
                            continue
                        except:
                            pass
                    elif element.get_attribute('id') == 'responsiveMainMenuBean-button-1':
                        eventDescription += 'Logged in successfully | '
                        break
                    else:
                        eventDescription += 'ERROR: Underlying web elements have been changed - dev required | '
            except:
                print(traceback.format_exc())
                eventDescription += 'ERROR: An issue occured whilst logging in or the server timed out (>20s)'
            finally:
                messageToEmit[HHHconf.events_EventDescription] = eventDescription
                self.signal_finished.emit(messageToEmit)

class equipmentFinanceEngine():

    # Note - some methods that INSERT into tables require that the sql table itself has ID column set to AUTO-INCREMENT by 1. That is, when inserting a new row, it auto-creates the column value.

    @staticmethod
    def getClientDetails(selectedClientList:list=None):
        if selectedClientList is None: # Get details of all Agreements
            job_getEFClientDetails = ('SELECT * FROM ' +  HHHconf.db_clientInfo_name)
        else: # Can selectively get clients with list of client numbers as input ['C00001', 'C00002'...]
            selectedClients = ['\'' +  x  + '\'' for x in selectedClientList] # adds raw apostrophes necessary for SQL query
            selectedClientsString = '(' + ', '.join(selectedClients) + ')'
            job_getEFClientDetails = ('SELECT * FROM ' +  HHHconf.db_clientInfo_name + ' WHERE ' + HHHconf.dict_toSQL_HHH_EF_Client_Info[HHHconf.client_number] + ' IN ' + selectedClientsString)
        df_clientDetails = pandas.read_sql(job_getEFClientDetails, mainEngine.establishSqlConnection())
        df_clientDetails.rename(columns=HHHconf.dict_fromSQL_HHH_EF_Client_Info, inplace=True)
        return df_clientDetails

    @staticmethod
    def getAgreementDetails(selectedAgreements:list=None):
        if selectedAgreements is None: # Get details of all Agreements
            job_getEFAgreementDetails = ('SELECT * FROM ' +  HHHconf.db_agreementInfo_name)
        else: # Get details of selected agreements  
            selectedAgreementsString = '(' + ', '.join(['\'' +  x  + '\'' for x in selectedAgreements]) + ')'
            job_getEFAgreementDetails = ('SELECT * FROM ' +  HHHconf.db_agreementInfo_name + ' WHERE ' + HHHconf.dict_toSQL_HHH_EF_Agreement_Info[HHHconf.agreement_number] + ' IN ' + selectedAgreementsString)
        df_agreementDetails = pandas.read_sql(job_getEFAgreementDetails, mainEngine.establishSqlConnection())
        df_agreementDetails.rename(columns=HHHconf.dict_fromSQL_HHH_EF_Agreement_Info, inplace=True)
        return df_agreementDetails

    @staticmethod
    def getPaymentDetails(selectedAgreements=None):
        if isinstance(selectedAgreements, list): # Get all payments in schedule for selectedAgreements list
            selectedAgreementsString = '(' + ', '.join(['\'' +  x  + '\'' for x in selectedAgreements]) + ')'
            job_getEFPaymentDetails = ('SELECT * FROM ' + HHHconf.db_payoutSchedule_name + ' WHERE ' + HHHconf.dict_toSQL_HHH_EF_Schedule[HHHconf.agreement_number] + ' IN ' + selectedAgreementsString)
        else:
            job_getEFPaymentDetails = ('SELECT * FROM ' + HHHconf.db_payoutSchedule_name)
        
        df_paymentDetails = pandas.read_sql(job_getEFPaymentDetails, mainEngine.establishSqlConnection())
        df_paymentDetails.rename(columns=HHHconf.dict_fromSQL_HHH_EF_Schedule, inplace=True)
        return df_paymentDetails

    @staticmethod
    def getCurrentSubordinates(selectedAgreement=None, selectedPayment=None, collapseGST=False, includeHistorical=False, backDate=None):
        if selectedAgreement is None:
            selectedAgreementCondition = ''
        else:
            selectedAgreementCondition = ' WHERE ' + HHHconf.dict_toSQL_HHH_EF_Subordinates[HHHconf.payment_number] + ' LIKE \'%' + str(selectedAgreement) + '%\''

        if selectedPayment is None:
            selectedPaymentCondition = ''
        else:
            selectedPaymentCondition = ' AND ' + HHHconf.dict_toSQL_HHH_EF_Subordinates[HHHconf.payment_number] + ' = \'' + str(selectedPayment) + '\''

        # Get all relevant subordinate transactions from db 
        job_getOutstandingSubordinates = ('SELECT * FROM ' + HHHconf.db_subordinate_name + selectedAgreementCondition + selectedPaymentCondition)
        with mainEngine.establishSqlConnection() as cnxn:
            df_currentSubordinates = pandas.read_sql(job_getOutstandingSubordinates, cnxn)
        df_currentSubordinates.rename(mapper=HHHconf.dict_fromSQL_HHH_EF_Subordinates, inplace=True, axis='columns')

        # Back-date the query (useful for client statements)
        if isinstance(backDate, datetime.datetime):
            df_currentSubordinates = df_currentSubordinates[df_currentSubordinates[HHHconf.subordinate_date] < backDate]
        
        # Drop duplicates to keep only the latest transaction for each subordinate        
        if includeHistorical is False:
            df_currentSubordinates = df_currentSubordinates.sort_values(by=['TransID'], ascending=False).drop_duplicates(subset=HHHconf.subordinate_number) 
            df_currentSubordinates.reset_index(drop=True, inplace=True)

        # collapse GST transactions into their parents
        if collapseGST is True:
            df_currentSubordinates[HHHconf.subordinate_amount] = df_currentSubordinates[HHHconf.subordinate_amount] + df_currentSubordinates[HHHconf.subordinate_gst]
            df_currentSubordinates[HHHconf.subordinate_gst] = 0

        return df_currentSubordinates

    @staticmethod
    def generateSubordinateTransactions(selectedAgreement=None, selectedPayment=None, collapseGST=False, includeHistorical=False, filterList=None, subordinateTypesList=None):
        df_currentSubordinates = equipmentFinanceEngine.getCurrentSubordinates(selectedAgreement=selectedAgreement, selectedPayment=selectedPayment, collapseGST=collapseGST)
        if filterList is not None:
            df_currentSubordinates = df_currentSubordinates[df_currentSubordinates[HHHconf.subordinate_status].isin(filterList)]
        if subordinateTypesList is not None:
            df_currentSubordinates = df_currentSubordinates[df_currentSubordinates[HHHconf.subordinate_status].isin(subordinateTypesList)]
        return df_currentSubordinates

    @staticmethod
    def createPayoutSchedulePreview(dictData):    # data requires fields from agreementInfo database
        numpy.set_printoptions(suppress=True, formatter={'float_kind':'{:0.6f}'.format}) # Required to suppress scientific notation in df later
        pandas.options.display.float_format = '{:0.6f}'.format
        # reformat data
        dict_dataFormatted = copy.deepcopy(dictData)
        dict_formatDataFunctions= {
            HHHconf.original_balance : float(dictData[HHHconf.original_balance]), 
            HHHconf.balloon_amount : float(dictData[HHHconf.balloon_amount]), 
            HHHconf.periodic_fee : float(dictData[HHHconf.periodic_fee]), 
            HHHconf.periods_per_year : int(dictData[HHHconf.periods_per_year]), 
            HHHconf.total_periods : int(dictData[HHHconf.total_periods]), 
            HHHconf.agreement_start_date : dictData[HHHconf.agreement_start_date] if isinstance(dictData[HHHconf.agreement_start_date], datetime.datetime) else datetime.datetime.strptime(dictData[HHHconf.agreement_start_date], HHHconf.dateFormat2), 
            HHHconf.settlement_date : dictData[HHHconf.settlement_date] if isinstance(dictData[HHHconf.settlement_date], datetime.datetime) else datetime.datetime.strptime(dictData[HHHconf.settlement_date], HHHconf.dateFormat2), 
            HHHconf.bsb : None if dictData[HHHconf.bsb] is None else str(dictData[HHHconf.bsb].replace('-', ''))
        }
        for key, value in dict_formatDataFunctions.items():
            try:
                dict_dataFormatted[key] = value
            except:
                pass
        
        # Calculate either Periodic Repayment OR Interest Rate
        try:
            if len(str(dictData[HHHconf.interest_rate])) > 0 and dictData[HHHconf.interest_rate] is not None:
                dict_dataFormatted[HHHconf.interest_rate] = float(dictData[HHHconf.interest_rate]) / 100
                dict_dataFormatted[HHHconf.periodic_repayment] = numpy.pmt(dict_dataFormatted[HHHconf.interest_rate]/dict_dataFormatted[HHHconf.periods_per_year], dict_dataFormatted[HHHconf.total_periods], - dict_dataFormatted[HHHconf.original_balance], dict_dataFormatted[HHHconf.balloon_amount], 'end')
            elif len(str(dictData[HHHconf.periodic_repayment])) > 0 and dictData[HHHconf.periodic_repayment] is not None:
                dict_dataFormatted[HHHconf.periodic_repayment] = float(dictData[HHHconf.periodic_repayment])
                dict_dataFormatted[HHHconf.interest_rate] = numpy.rate(dict_dataFormatted[HHHconf.total_periods], - dict_dataFormatted[HHHconf.periodic_repayment], dict_dataFormatted[HHHconf.original_balance], dict_dataFormatted[HHHconf.balloon_amount]) * dict_dataFormatted[HHHconf.periods_per_year]
        except Exception as e:
            print('ERROR: ' + str(e))
            
        # Create payout schedule
        list_payoutSchedule = []
        dict_dateIncrementer = {
            4: (lambda x: dict_dataFormatted[HHHconf.agreement_start_date] + relativedelta(months=3*x)), 
            12: (lambda x: dict_dataFormatted[HHHconf.agreement_start_date] + relativedelta(months=x)), 
            26: (lambda x: dict_dataFormatted[HHHconf.agreement_start_date] + datetime.timedelta(days=14*x)), 
            52: (lambda x: dict_dataFormatted[HHHconf.agreement_start_date] + datetime.timedelta(days=7*x))
        }
        for i in range(0, dict_dataFormatted[HHHconf.total_periods]):
            currentPeriodData = {}
            currentPeriod = i + 1
            currentPeriod_openingBalance = dict_dataFormatted[HHHconf.original_balance] if currentPeriod == 1 else previousPeriod_closingBalance
            try:
                currentPeriod_payoutDate = dict_dateIncrementer[dict_dataFormatted[HHHconf.periods_per_year]](currentPeriod - 1)
            except:
                currentPeriod_payoutDate = 'not defined'
            currentPeriod_repaymentAmount = dict_dataFormatted[HHHconf.periodic_repayment]
            currentPeriod_interestAmount = numpy.round(numpy.ipmt(dict_dataFormatted[HHHconf.interest_rate]/dict_dataFormatted[HHHconf.periods_per_year], currentPeriod, dict_dataFormatted[HHHconf.total_periods], - dict_dataFormatted[HHHconf.original_balance], dict_dataFormatted[HHHconf.balloon_amount], 'end'), decimals=10)
            currentPeriod_principalAmount = float(currentPeriod_repaymentAmount - currentPeriod_interestAmount)
            currentPeriod_closingBalance = float(currentPeriod_openingBalance + currentPeriod_interestAmount - currentPeriod_repaymentAmount)
            currentPeriodData = {
                'currentPeriod': currentPeriod, 
                HHHconf.repayment_date: currentPeriod_payoutDate, 
                HHHconf.opening_balance: currentPeriod_openingBalance, 
                HHHconf.interest_component: currentPeriod_interestAmount, 
                HHHconf.principal_component: currentPeriod_principalAmount, 
                HHHconf.periodic_repayment: currentPeriod_repaymentAmount, 
                HHHconf.closing_balance: currentPeriod_closingBalance
            }
            list_payoutSchedule.append(currentPeriodData)
            previousPeriod_closingBalance = currentPeriod_closingBalance
        
        # Create dataframe
        df_payoutSchedule = pandas.DataFrame(list_payoutSchedule)
        df_payoutSchedule = df_payoutSchedule[['currentPeriod', HHHconf.repayment_date, HHHconf.opening_balance, HHHconf.interest_component, HHHconf.principal_component, HHHconf.periodic_repayment, HHHconf.closing_balance]] # re-order columns
        return dict_dataFormatted, df_payoutSchedule, dict_dataFormatted[HHHconf.periodic_repayment]

        
    @staticmethod
    def createNewClient(dict_clientDetails): 
    # data requires fields from clientInfo database
        with mainEngine.establishSqlConnection() as cnxn:
            cursor = cnxn.cursor()
            columnNames = ', '.join([HHHconf.dict_toSQL_HHH_EF_Client_Info[x] for x in list(dict_clientDetails.keys())])
            placeholders = ', '.join(['?'] * len(dict_clientDetails))
            Job_addEFClient = ('INSERT INTO ' + HHHconf.db_clientInfo_name + ' (' + columnNames + ') OUTPUT INSERTED.ID VALUES (' + placeholders + ')')
            cursor.execute(Job_addEFClient, list(dict_clientDetails.values()))
            newClientID = cursor.fetchone()
            cnxn.commit()      
            # Generate Client Number
            newClientNumber = 'C{0:0>5}'.format(str(newClientID[0]))   
            cursor.execute('UPDATE ' + HHHconf.db_clientInfo_name + ' SET ' + HHHconf.dict_toSQL_HHH_EF_Client_Info[HHHconf.client_number] + ' = ? WHERE ID = ' + str(newClientID[0]), newClientNumber)
            cnxn.commit()
        return newClientNumber

    @staticmethod
    def createAgreement(dict_agreementDetails):
        with mainEngine.establishSqlConnection() as cnxn:
            cursor = cnxn.cursor()
            columnNames = ', '.join([HHHconf.dict_toSQL_HHH_EF_Agreement_Info[x] for x in list(dict_agreementDetails.keys())])
            placeholders = ', '.join(['?'] * len(dict_agreementDetails))
            Job_addEFAgreement = ('INSERT INTO ' + HHHconf.db_agreementInfo_name + ' (' + columnNames + ') OUTPUT INSERTED.ID VALUES (' + placeholders + ')')
            cursor.execute(Job_addEFAgreement, list(dict_agreementDetails.values()))
            newAgreementID = cursor.fetchone()
            cnxn.commit()      
            # Generate Agreement Number
            newAgreementNumber = 'EF{0:0>7}'.format(str(newAgreementID[0]))   
            cursor.execute('UPDATE ' + HHHconf.db_agreementInfo_name + ' SET ' + HHHconf.dict_toSQL_HHH_EF_Agreement_Info[HHHconf.agreement_number] + ' = ? WHERE ID = ' + str(newAgreementID[0]), newAgreementNumber)
            cnxn.commit()
        return newAgreementNumber

    @staticmethod
    def commitDataframeToSQLDB(df, dbName):
        engine = create_engine('mssql+pymssql://' + HHHconf.sqlUsername + ':' + HHHconf.sqlPassword + '@' + HHHconf.sqlServer + '/' + HHHconf.sqlDatabase)
        df.to_sql(name=dbName, index=False, if_exists='append', con=engine)
        del df
        engine.dispose()

    # TODO: fully automate EF start of day process 
    @QtCore.pyqtSlot()
    def startOfDay():
        print('test bod')
        # Update subordinate statuses based on today's date
        df_subordinates = equipmentFinanceEngine.generateSubordinateTransactions(filterList=[HHHconf.dict_paymentStatus[HHHconf.paymentStatus_pending], HHHconf.dict_paymentStatus[HHHconf.paymentStatus_rescheduled], HHHconf.dict_paymentStatus[HHHconf.paymentStatus_due], HHHconf.dict_paymentStatus[HHHconf.paymentStatus_overdue]])
            # from this dataframe - any transactions dated today, update to due; any transactions dated older than today, update to overdue
        datetimeToday = datetime.datetime.combine(datetime.date.today(), datetime.time())
        df_subordinatesToUpdate = copy.deepcopy(df_subordinates.loc[df_subordinates[HHHconf.subordinate_value_date] <= datetimeToday])
        datetimeNow = datetime.datetime.now()
            # Also update subordinate datetime to now
        df_subordinatesToUpdate.loc[(df_subordinatesToUpdate[HHHconf.subordinate_value_date] < datetimeToday) & (df_subordinatesToUpdate[HHHconf.subordinate_status] != HHHconf.dict_paymentStatus[HHHconf.paymentStatus_overdue]), [HHHconf.subordinate_status, HHHconf.subordinate_date]] = HHHconf.dict_paymentStatus[HHHconf.paymentStatus_overdue], datetimeNow
        df_subordinatesToUpdate.loc[(df_subordinatesToUpdate[HHHconf.subordinate_value_date] == datetimeToday) & (df_subordinatesToUpdate[HHHconf.subordinate_status] != HHHconf.dict_paymentStatus[HHHconf.paymentStatus_due]), [HHHconf.subordinate_status, HHHconf.subordinate_date]] = HHHconf.dict_paymentStatus[HHHconf.paymentStatus_due], datetimeNow
            # update notes to system generated
        df_subordinatesToUpdate[HHHconf.subordinate_notes] = HHHconf.subordinateNotes_systemGenerated
        df_subordinatesToUpdate[HHHconf.subordinate_created_by] = HHHconf.name_app
        
        df_subordinatesToUpdate.drop(columns=['TransID'], inplace=True)
        df_subordinatesToUpdate.drop(df_subordinatesToUpdate[df_subordinatesToUpdate[HHHconf.subordinate_date] != datetimeNow].index, inplace=True) # Only keep subordinates that were updated 
            # add transactions to db to update
        df_subordinatesToUpdate.rename(mapper=HHHconf.dict_toSQL_HHH_EF_Subordinates, inplace=True, axis='columns')
        equipmentFinanceEngine.commitDataframeToSQLDB(df_subordinatesToUpdate, HHHconf.db_subordinate_name)

        print(df_subordinatesToUpdate)
        
        # THIS IS OLD CODE - TODO: clean up or repurpose in the event it is required
        # WE DONT REQUIRE WPC DISHONOUR ANYMORE - MPOWER WILL AUTO PICK UP DISHONOURS AND UPDATE THE STATUS FROM UNCLEARED TO OVERDUE
        # # Access Dishonour XL at beginning of day into DF
        # xmlFile = 'C:\\Users\\Hugh.huang\\Desktop\\dishonour.xml'
        # tree = et.parse(open(xmlFile))
    
        # list_ColumnNames = []
        # listDishonours = []
        # for index, dishonour in enumerate(tree.iter('PaymentBatchEntry')):
        #     list_currentDishonour = []
        #     for detail in dishonour.iter():
        #         if len(list(detail)) == 0:
        #             if index == 0:
        #                 list_ColumnNames.append(str(detail.tag))
        #             list_currentDishonour.append(str(detail.text))
        #     listDishonours.append(list_currentDishonour)
        # df_dishonours = pandas.DataFrame(listDishonours, columns=list_ColumnNames)
        # # pandas.set_option('display.max_columns', None)
        # # print(df_dishonours)

        # # Filter down DF to relevant data
        # df_dishonoursFiltered = df_dishonours[['BatchEntryID', 'BankReference', 'BuyerCompanyID', 'BuyerCompanyName', 'AccountName', 'PaymentAmount', 'PrincipalAmount', 'SurchargeAmount', 'CreatedDate', 'PaymentDate']]
        # print(df_dishonoursFiltered)
        # # Associate Mtx accounts with EF Agreements
        # # Update database with new payment statuses

    # TODO: automate EF end of day process
    @QtCore.pyqtSlot()
    def endOfDay():
        print('test EOD')        
        # Check if beginning of day has already been done - if not, return
        # Send due transactions to mPower
        # include agreement number column, and collate by same agreement + same payment type 
        df_dueTransactions = equipmentFinanceEngine.generateSubordinateTransactions(filterList=[HHHconf.dict_paymentStatus[HHHconf.paymentStatus_due]])
        # Create list here, before columns are renamed
        list_dueTransactions = df_dueTransactions.to_dict('records')
        if len(list_dueTransactions) == 0:
            return
        
        # Update these transaction statuses to UNCLEARED, which will lock them from users until they are settled
        df_dueTransactions[HHHconf.subordinate_status] = HHHconf.dict_paymentStatus[HHHconf.paymentStatus_uncleared]
        df_dueTransactions[HHHconf.subordinate_date] = datetime.datetime.now()
        df_dueTransactions[HHHconf.subordinate_notes] = HHHconf.subordinateNotes_systemGenerated
        df_dueTransactions[HHHconf.subordinate_created_by] = HHHconf.name_app
        df_dueTransactions.drop(columns=['TransID'], inplace=True)
        df_dueTransactions.rename(mapper=HHHconf.dict_toSQL_HHH_EF_Subordinates, inplace=True, axis='columns')
        equipmentFinanceEngine.commitDataframeToSQLDB(df_dueTransactions, HHHconf.db_subordinate_name)

        # Temporarily gather P + I components before joining with due transactions
        list_principalAndInterestComponents = []
        for dict_row in list_dueTransactions:
            if dict_row[HHHconf.subordinate_type] == HHHconf.dict_subordinateType[HHHconf.subordinateType_repayment]:
                df_paymentDetails = equipmentFinanceEngine.getPaymentDetails(selectedAgreements=[dict_row[HHHconf.payment_number].split('_')[0]])
                # Get subordinate P + I details from payout schedule
                dict_paymentDetails = df_paymentDetails[df_paymentDetails[HHHconf.payment_number] == dict_row[HHHconf.payment_number]].to_dict('records')[0]
                # Calculate Interest and Principal components to TAKE - this would differ to actual P + I recorded in the table due to complex rounding issues
                dict_principalTake = dict.fromkeys(dict_row, None)
                for x in list(HHHconf.dict_toSQL_HHH_EF_Subordinates.keys()):
                    if x == HHHconf.subordinate_type:
                        dict_principalTake[x] = str(HHHconf.dict_toSQL_HHH_EF_Schedule[HHHconf.principal_component])
                    elif x == HHHconf.subordinate_amount:
                        dict_principalTake[x] = float(round(dict_paymentDetails[HHHconf.opening_balance], 2) - round(dict_paymentDetails[HHHconf.closing_balance], 2)) # calculates to the cent, what principal to take to get the right closing balance, in line with Mtx  
                    else:
                        dict_principalTake[x] = dict_row[x]
                dict_interestTake = dict.fromkeys(dict_row, None)
                for x in list(HHHconf.dict_toSQL_HHH_EF_Subordinates.keys()):
                    if x == HHHconf.subordinate_type:
                        dict_interestTake[x] = str(HHHconf.dict_toSQL_HHH_EF_Schedule[HHHconf.interest_component])
                    elif x == HHHconf.subordinate_amount:
                        dict_interestTake[x] = float(dict_row[x] - dict_principalTake[x]) # allocates remainder of repayment as interest (after deducting principal)
                    else:
                        dict_interestTake[x] = dict_row[x]
                list_principalAndInterestComponents.append(dict_principalTake)
                list_principalAndInterestComponents.append(dict_interestTake)
        # join P + I components, then remove any periodic repayment transactions (as they are covered by the new P + I transactions)
        list_dueTransactions  += list_principalAndInterestComponents
        list_dueTransactions = [x for x in list_dueTransactions if HHHconf.dict_subordinateType[HHHconf.subordinateType_repayment] not in x[HHHconf.subordinate_type]]

        # Convert back to dataframe, then configure df necessary for mPower EF Instructions table
        df_dueTransactions = pandas.DataFrame(list_dueTransactions)
        print(df_dueTransactions)
        df_eod = pandas.DataFrame()
        df_eod[HHHconf.EFInstructions_valueDate] = df_dueTransactions[HHHconf.subordinate_value_date]
        df_eod[HHHconf.EFInstructions_agreementNumber] = df_dueTransactions[HHHconf.payment_number].str.split('_', expand=True)[0]
        df_eod[HHHconf.EFInstructions_transactionType] = df_dueTransactions[HHHconf.subordinate_type]
        df_eod[HHHconf.EFInstructions_status] = 'Pending'
        df_eod[HHHconf.EFInstructions_amount] = df_dueTransactions[HHHconf.subordinate_amount] + df_dueTransactions[HHHconf.subordinate_gst]

        print(df_eod) # Here's where the transactions should get sent to Instructions table

    @staticmethod
    def addGSTtoTransactions(transactions, transactionTypeKey, transactionAmountKey, transactionGSTKey):
        # returns df with gst added to relevant transaction
        if isinstance(transactions, pandas.DataFrame):
            transactions[transactionGSTKey] = 0
            transactions.loc[transactions[transactionTypeKey].isin(HHHconf.list_subordinateGST), transactionGSTKey] = transactions.loc[transactions[transactionTypeKey].isin(HHHconf.list_subordinateGST), transactionAmountKey] * HHHconf.gstComponent
        return transactions
    
    @staticmethod
    def terminateAgreement(selectedAgreements:list, terminationDateTime=None, terminationNotes:str=HHHconf.subordinateNotes_systemGenerated):
        # TODO
        # Set agreement status to closed
        # Set all pending payments in the schedule to Complete, completionDate = terminationDate, concat payment notes with termination notes
        # Set all outstanding subordinates to credited, subordinateDate = terminationDate, Subordinate notes = termination notes, created by HHH
        # Output all subordinate transactions that were credited in the process - need a termination report with the changes done before closing the account

        # FOR TESTING - Update EF0000002 to Active status, then try terminating that agreement (testClient):
        # UPDATE HHH_EF_Agreement_Info SET AgreementStatus = 'Closed' WHERE EFAgreementNumber = 'EF0000002'
        return
