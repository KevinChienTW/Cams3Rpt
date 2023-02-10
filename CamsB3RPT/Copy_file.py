import shutil
import openpyxl as xl
import os 
from copy import copy
from openpyxl.worksheet.datavalidation import DataValidation
import time
import openpyxl as x1



class Copy_report:
    def copy_report(DUTtype):
        report = 'report.xlsx'
        file = 'D:/python/230209/xls/Templates.xlsx'
        shutil.copyfile(file,report)  

                                  
    def main():
        report = 'report.xlsx'
        file = 'D:/python/230209/xls/Templates.xlsx'
        shutil.copyfile(file,report)  
        # kevin add -s
        wb = xl.load_workbook('report.xlsx')
        #ws = wb['test']
        ws = wb.worksheets[0]
        ws.cell(row = 1 ,column = 1).value = 'kkkkkkk'
        wb.save('report.xlsx')


        path = 'D:/python/230209/summary_1.txt'
        f = open(path, 'r')
        print(f.read())
        f.close()


        f = open(path, "a")
        f.write( "123\n")
        f.close()

        # 使用r+
        f = open(path, "r+")
        f.write( "456\n")
        f.close()

        # 使用a+
        f = open(path, "a+")
        f.write( "789\n")
        f.close()

        # kevin add -E

    
Copy_report.main()

