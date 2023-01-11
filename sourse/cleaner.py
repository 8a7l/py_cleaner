#===================================================
# Підключенні бібліотеки.
#===================================================
import openpyxl
import os
import time
#===================================================


#===================================================
# Головні змінні.
#===================================================
file_way=os.path.dirname(__file__)
file_name=os.path.basename(__file__)
file_name=file_name.split('.')
file_name=file_name[0]
database_excel=file_name+'.xlsx'
script_file=file_name+'.bat'
script_file_patern=R'''@echo off
del /s /f /q c:\windows\temp\*.*
rd /s /q c:\windows\temp
md c:\windows\temp
del /s /f /q C:WINDOWS\Prefetch
del /s /f /q %temp%\*.*
rd /s /q %temp%
md %temp%
deltree /y c:\windows\tempor~1
deltree /y c:\windows\temp
deltree /y c:\windows\tmp
deltree /y c:\windows\ff*.tmp
deltree /y c:\windows\prefetch
deltree /y c:\windows\history
deltree /y c:\windows\cookies
deltree /y c:\windows\recent
deltree /y c:\windows\spool\printers
cls
exit'''
#===================================================


#===================================================
# Функції
#===================================================
def database_excel_open():
    workbook = openpyxl.load_workbook(database_excel)
    workbook_sheet=workbook['sheet1']
    a1=workbook_sheet['A1'].value
    f = open(script_file, 'r').read()
    if hash(f) == hash(script_file_patern):
        os.system('start '+script_file)
        print(time.asctime())
        time.sleep(a1)
    else:
        database_excel_create()



def database_excel_create():
    workbook = openpyxl.Workbook()
    workbook_sheet=workbook.active
    workbook_sheet.title = 'sheet1'
    workbook_sheet['A1'] = 300
    workbook_sheet['B1'] = 'Час через який буде повторно запущено файл.(в секундах)'
    workbook_sheet['B2'] = 900
    workbook_sheet['B3'] = 1800
    workbook_sheet['B4'] = 3600
    workbook_sheet['C2'] = '15 хвилин'
    workbook_sheet['C3'] = '30 хвилин'
    workbook_sheet['C4'] = '1 година'
    workbook.save(database_excel)
    os.remove(script_file)
    f = open(script_file, 'a+')
    f.write(script_file_patern)
    f.close()


def fail_message():
    print('Виявлена помилка.')
    print('Всі налаштування скинуто до початкових.')
    print('Перезапустіть програму.')
    input()

#===================================================


#===================================================
# Точка входу.
#===================================================
def main():
    try:
        while True:
            database_excel_open()
    except:
        database_excel_create()
        fail_message()


if __name__ == '__main__':
    main()
#===================================================
