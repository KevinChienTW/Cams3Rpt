import os
import openpyxl as xl
import tkinter
from xlrd import open_workbook
from tkinter import messagebox
from win32com.client import Dispatch



class Copy_report:
    def copy_only_C(category,DUTtype,file,report):
        xlap = Dispatch("Excel.Application")      
        xlap.Visible = True
        wb1 = xlap.Workbooks.Open(Filename = file)      #open template file
        wb2 = xlap.Workbooks.Open(Filename = report)    #open report
        ws1 = wb1.Worksheets('USB_Detail')
        ws1.Copy(Before = wb2.Worksheets(len(wb2.Worksheets)))
        ws1 = wb1.Worksheets('Spec.')
        ws1.Copy(Before = wb2.Worksheets(len(wb2.Worksheets)))
        wb2.Close(SaveChanges=True)
        wb1.Close()


    def copy_report_sheet(template,DUTtype,DUTcategory,DUTUP,DUTMaxSpeed):
        print(template)
        new_report = xl.Workbook()
        new_report.save('report.xlsx')
        xlap = Dispatch("Excel.Application")      
        xlap.Visible = False
        report = os.getcwd() + '/report.xlsx'
        wb1 = xlap.Workbooks.Open(Filename = report)      #open template file
        wb2 = xlap.Workbooks.Open(Filename = template)    #open report
        if DUTUP == 'Type-C':
            if DUTcategory == 'Device':
                a=1
            elif DUTcategory == 'Hub':
                a=1
            elif DUTcategory == 'Device':
                a=1
            elif DUTcategory == 'Host':
                a=1
            elif DUTcategory == 'Embedded-Host':  
                a=1
           
                                  
    def copy_template(DUTtype,DUTcategory,DUTUP,DUTMaxSpeed):
        if DUTUP == 'Type-C':
            path = './allion_usb_template_D-03172022-221007/TypeC (USB&PD)'
        else:
            path = './allion_usb_template_D-03172022-221007/USB only (Legacy)'

        files = os.listdir(path)
        template = ''
        for i in files:
            if DUTtype in i and DUTcategory in i and DUTMaxSpeed in i:
                template =  os.getcwd() + path + '/' + i
                Copy_report.copy_report_sheet(template,DUTtype,DUTcategory,DUTUP,DUTMaxSpeed) 

            elif DUTcategory == 'Embedded-Host' and DUTMaxSpeed in i:
                template =  os.getcwd() + path + '/' + i
                Copy_report.copy_report_sheet(template,DUTtype,DUTcategory,DUTUP,DUTMaxSpeed) 
            
            elif DUTMaxSpeed == 'USB2.0' and DUTtype in i and DUTcategory in i and 'Gen1' in i:
                template =  os.getcwd() + path + '/' + i
                Copy_report.copy_report_sheet(template,DUTtype,DUTcategory,DUTUP,DUTMaxSpeed) 
            
            elif DUTcategory == 'Embedded-Host' and DUTMaxSpeed == 'USB2.0' and 'Gen1' in i:
                template =  os.getcwd() + path + '/' + i
                Copy_report.copy_report_sheet(template,DUTtype,DUTcategory,DUTUP,DUTMaxSpeed) 
        return template
         
                                  
    def main(DUTtype,DUTcategory,DUTUP,DUTMaxSpeed,controller,DC,DA,DH):

        try:
            template = Copy_report.copy_template(DUTtype,DUTcategory,DUTUP,DUTMaxSpeed)
            Copy_report.copy_report_sheet(template,DUTtype,DUTcategory,DUTUP,DUTMaxSpeed)
            result = 'Pass' 

            return result
        except PermissionError:
            result = 'Fail' 
            print('Copy error,please colse report.xlsx')
            return result

Copy_report.main('End-Product','Device','Type-C','USB2.0',1,1,1,1)