
from os import listdir
from os.path import isfile, join
import glob
import os
from openpyxl import Workbook
from openpyxl import load_workbook


def rename_file(path):
    dir = glob.glob(path+"/*.xls")
    for old_file in dir:
        old_filename = old_file.split('\\')[-1].split('.xls')[0]
        new_file = join(path,old_filename+".html")
        os.rename(old_file,new_file)

        
def save_file_as_xlsx(path):
    dir = glob.glob(path+"/*.xls")
    for old_file in dir:
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        wb = excel.Workbooks.Open(old_file)
        wb.SaveAs(old_file+"x", FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
        wb.Close()                               #FileFormat = 56 is for .xls extension
        excel.Application.Quit()        
        
        
def rename_worksheet(path):
    dir = glob.glob(path+"/*.xlsx")
    for old_file in dir:
        wb = load_workbook(old_file)
        ws =  wb.active
        ws.title = "TargetedArea"
        wb.save(filename = old_file)       
        
        
#rename_file(path)
#save_file_as_xlsx(path)
#rename_worksheet(path)
