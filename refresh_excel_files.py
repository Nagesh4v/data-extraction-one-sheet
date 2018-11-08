import csv
import os
import win32com.client
import datetime
import openpyxl

def refresh_cc_sheets(file_list):
# Open Excel
    #check if file is in read only mode
    
    Application = win32com.client.Dispatch("Excel.Application")
 
 # Show Excel. While this is not required, it can help with debugging
    Application.Visible = 1
    Application.DisplayAlerts=False
    Application.AskToUpdateLinks = False
    #Application.Calculation = -4105

    for  full_path in file_list:
        print(full_path)
        #Application.Calculation = -4135
        #Application.Calculation = -4105
        # Open Your Workbook
        try:
            #Workbook = Application.Workbooks.open(path + input_file + country + '.xlsx')
            fd = os.open(full_path,os.O_RDWR)
            os.close(fd)
            Workbook = Application.Workbooks.open(full_path)
        except Exception as e:
            print('Workbook already opened!')
        else:
            try:
                Workbook.UpdateLink(Name=Workbook.LinkSources())

            except Exception as e:
                print(e)
            # Refesh All
            Workbook.RefreshAll()
            Application.CalculateUntilAsyncQueriesDone()
            Application.Calculate()
            # Saves the Workbook
            Workbook.Save()
            Workbook.Close()
            print(full_path + ' - Done')
            #print('{}{} Done Started at:{} Ended at:{}'.format(full_path,started,datetime.datetime.now().strftime("%H:%M")))


    Application.Visible = 0
    Application.DisplayAlerts=True
    Application.AskToUpdateLinks = True
 # Closes Excel
    Application.Quit()

now=datetime.datetime.now()
cur_date=now.strftime("%Y%m%d")
print('Started at: {}'.format(now.strftime("%H:%M")))
file_refresher='Z:\\600-Marketing\\685-SEM\\686-Reports\\BI Queries\\File refresher.xlsm'
#open the workbook
started=datetime.datetime.now().strftime("%H:%M")

wb = openpyxl.load_workbook(file_refresher,read_only=True, data_only=True)
ws = wb['Admin']
file_list=[]
for row in ws.iter_rows(min_row=2,max_row=10, min_col=2, max_col=2):
    if row[0].value != None:
        file_list.append(row[0].value)
print(file_list)
wb.close
refresh_cc_sheets(file_list)