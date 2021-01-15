#Import libraries
import pandas as pd
import PySimpleGUI as sg
from datetime import datetime
#Function to change column headers by adding a small value to the question ID
def col_change(df, num):
    new_cols = []
    cols = list(df.columns)
    for col in cols:
        new_cols.append(col + num)
    df.columns = new_cols
    return df

#Create GUI to input file path and type of survey

sg.theme('DarkAmber')
# All the stuff inside your window.
layout = [[sg.Text('Select File:')],
          [sg.In(),sg.FileBrowse()],
          [sg.Text('Select Survey:')],
          [sg.Combo(['About Me', 'About My Dog', ' About My Household', '3 Week', '2.5 Month'])],
          [sg.Button('Ok'), sg.Button('Cancel')]]

# Create the Window
window = sg.Window('Window Title', layout)
# Event Loop to process "events" and get the "values" of the inputs
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Cancel': # if user closes window or clicks cancel
        break
    if event == 'Ok': #if user clicks okay
        file_path = values.get(0)
        survey = values.get(1)
window.close()

#Get current datetime
now = datetime.now()
dt_string = now.strftime("%Y%m%d%H%S")

#Read Excel Inputs
'''Survey'''
df = pd.read_excel(file_path)
'''Allias'''


#Pivot for primary question response and secondary question response
pvt_pa = df.pivot(index='Registration ID', columns='Question Id', values='Question Response Answer')
pvt_sa = df.pivot(index='Registration ID', columns='Question Id', values='Question Response Secondary Answer')
#Change column names for secondary answers
pvt_sa = col_change(pvt_sa, 0.1)
#Drop empty columns for secondary answers
pvt_sa_cln = pvt_sa.dropna(axis=1, how='all')
#Join primary and secondary response dataframes
pvt_join = pd.concat([pvt_pa, pvt_sa_cln], axis=1, join="inner")
#Sort columns
pvt_final = pvt_join.reindex(sorted(pvt_join.columns), axis=1)
#Write to excel output
pvt_final.to_excel(r'C:\Python\TestProject\About_Me_Test_Output_pivot_{}.xlsx'.format(dt_string))
print("Pivot Complete")

