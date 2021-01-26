#Import libraries
import pandas as pd
import PySimpleGUI as sg
from datetime import datetime
import sys
#Function to change column headers by adding a small value to the question ID
def col_change(df, num):
    new_cols = []
    cols = list(df.columns)
    for col in cols:
        new_cols.append(col + num)
    df.columns = new_cols
    return df

#Function to return the correct tab in the reference doc
def survey_select(surv):
    if surv == 'About Me':
        lookup = 0
    elif surv == 'About My Dog':
        lookup = 1
    elif surv == 'About My Household':
        lookup = 2
    elif surv == 'Leaving Study':
        lookup = 3
    elif surv == '3 Week':
        lookup = 4
    elif surv == '2.5 Month':
        lookup = 5
    return lookup

#Create GUI to input file path and type of survey
sg.theme('DarkAmber')
# All the stuff inside window.
layout = [[sg.Text('Select File:')],
          [sg.In(),sg.FileBrowse()],
          [sg.Text('Select Survey:')],
          [sg.Combo(['About Me', 'About My Dog', 'About My Household', 'Leaving Study', '3 Week', '2.5 Month'])],
          [sg.Button('Ok'), sg.Button('Cancel')]]

# Create the Window
window = sg.Window('Window Title', layout)
# Event Loop to process "events" and get the "values" of the inputs
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Cancel': # if user closes window or clicks cancel
        window.close()
        sys.exit()
        break
    if event == 'Ok': #if user clicks okay
        file_path = values.get(0)
        survey = values.get(1)
        window.close()
        break


#Get current datetime
now = datetime.now()
dt_string = now.strftime("%Y%m%d%H%S")



#Read Excel Inputs
'''Survey'''
df = pd.read_excel(file_path)
'''Alias'''
df_ref = pd.read_excel('C:\Python\TestProject\QuestionRef.xlsx', survey_select(survey))

'''Pivoting and cleaning'''
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
pvt_sort = pvt_join.reindex(sorted(pvt_join.columns), axis=1)

'''Aliasing'''
#Clean Reference Table
df_ref_cln = df_ref.dropna(axis=0, how='all')

#Create Reference Dictionary
change_dict = {}
for value in df_ref_cln.values:
    change_dict[value[0]] = value[2]
    change_dict[value[0] + 0.1] = (str(value[2]) + ' (Further Info)')

#Rename columns
pvt_final = pvt_sort.rename(columns=change_dict, inplace=False)

#Write to excel output
pvt_final.to_excel(r'C:\Python\TestProject\Output\{}_pivot_{}.xlsx'.format(survey, dt_string))
print("Pivot Complete")

