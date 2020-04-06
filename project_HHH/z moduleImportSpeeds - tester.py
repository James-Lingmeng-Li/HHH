import timeit

for mod in ['sys', 'os', 'math', 'subprocess', 'pandas'',' 'datetime', 'time', 'calendar', 'numpy', 'pyodbc', 're', 'openpyxl', 'docx', 'shutil', 'threading', 'traceback']:
    x = timeit.timeit('import ' + mod)
    print('imported ' +mod+ ' in %s seconds'%x)

x = timeit.timeit('from PyQt5 import QtCore, QtGui, QtWidgets')
print('imported PyQT5 in %s seconds'%x)

x = timeit.timeit('import HHHconf, HHHfunc, class_EFAddAgreementScreen, class_EFViewAgreementScreen, class_EFAmendPaymentScreen, class_EFViewClientScreen, class_settingsScreen, class_mailMergerScreen, class_DFBankingScreen, class_DFB2BScreen')
print('imported HHH modules in %s seconds'%x)

print('done')