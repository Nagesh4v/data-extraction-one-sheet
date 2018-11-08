import csv
import os
import win32com.client
import datetime

def refresh_cc_sheets(path,input_file,countries):
# Open Excel
    full_path=path + input_file + '.xlsx'
    #check if file is in read only mode
    
    Application = win32com.client.Dispatch("Excel.Application")
 
 # Show Excel. While this is not required, it can help with debugging
    Application.Visible = 1
    Application.DisplayAlerts=False
    Application.AskToUpdateLinks = False
    #Application.Calculation = -4105

    for  country in countries.values():
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
            print(country + ' - Done')


    Application.Visible = 0
    Application.DisplayAlerts=True
    Application.AskToUpdateLinks = True
 # Closes Excel
    Application.Quit()

now=datetime.datetime.now()
cur_date=now.strftime("%Y%m%d")
countries={2:"UK", 9:"Austria", 3:"Europe", 6:"France", 1:"Germany", 11:"Italy", 13:"Spain", 7:"Switzerland",  12:"USA", 14:"Netherlands", 10:"Belgium"}
path='Z:\\800-Management\\830-Controlling\\833-Marketing\\Channel Controlling 2018\\'
input_file='Channel Controlling 2018 '
print('Started at: {}'.format(now.strftime("%H:%M")))

refresh_cc_sheets(path,input_file,countries)