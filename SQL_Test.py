#Import libriaries
import pyodbc
import PySimpleGUI as sg
import sys
import pandas as pd
from datetime import datetime

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
    elif surv == 'About My Household':
        lookup = 1
    elif surv == 'About {{DOG_NAME}}':
        lookup = 2
    #elif surv == 'About My Dog (Pre 25.11.2020)':
    #    lookup = 3
    elif surv == 'I want to remove {{DOG_NAME}} from the study':
        lookup = 3
    elif surv == '3 Week Survey':
        lookup = 4
    elif surv == '2.5 Month Survey':
        lookup = 5
    elif surv == '6 Month Survey':
        lookup = 6
    return lookup

#Create GUI to input file path and type of survey
sg.theme('DarkAmber')
# All the stuff inside window.
layout = [[sg.Text('Select Survey:')],
          [sg.Combo(['About Me', 'About My Household', 'About {{DOG_NAME}}', 'I want to remove {{DOG_NAME}} from the study', '3 Week Survey', '2.5 Month Survey', '6 Month Survey'])],
          [sg.Text('Select Output Folder:')],
          [sg.In(),sg.FolderBrowse()],
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
        survey = values.get(0)
        output_path = values.get(1)
        window.close()
        break

#SQL Server
conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=dtstrayapp01;'
                      'Database=paws;'
                      'Trusted_Connection=yes;')
#Create cursor object
cursor = conn.cursor()

df = pd.read_sql_query(
    """
    select
       a.registration_id "Registration ID",
       u.username "URN",
       d.dts_name "DTS Dog Name",
       d.animal_id "DTS Animal Id",
       s.name "Survey Name",
       ss.order_num "Section Num",
       ss.name "Section Name",
       sss.order_num "Sub Section Num",
       sss.name "Sub Section Name",
       sq.id "Question Id",
       sq.text "Question Text",
       sq.type "Question Type",
       sq.required "Question Required Flag",
       sq.dependent_on_question_answer "Dependent Question Required Answer",
       qr.id "Question Response Id",
       qr.answer "Question Response Answer",
       qr.secondary_answer "Question Response Secondary Answer",
       dqr.answer "Dependent Question Actual Answer"
    from survey s

       inner join survey_response sr on
              sr.survey_id = s.id
       inner join dog d on
              d.id = sr.dog_id
       inner join [user] u on
              u.id = sr.user_id
       inner join dbo.adoption a on
        a.[user_id] = u.id and d.id = a.dog_id
       inner join survey_section ss on
              s.id = ss.survey_id
       inner join survey_sub_section sss on
              ss.id = sss.survey_section_id
       inner join survey_question sq on
              sq.survey_sub_section_id = sss.id
       left outer join question_response qr on
              qr.survey_question_id = sq.id
              and qr.survey_response_id = sr.id
       left outer join survey_question dq on
              dq.id = sq.dependent_on_question_id
       left outer join question_response dqr on
              dqr.survey_question_id = dq.id and
              dqr.survey_response_id = sr.id
where s.name = N'{}'
      and qr.answer is not NULL 
      and qr.answer != ' '
order by
       u.username,
       d.animal_id,
       ss.order_num,
       sss.order_num,
       sq.order_num;""".format(survey) ,conn)

#Get current datetime
now = datetime.now()
dt_string = now.strftime("%Y%m%d%H%S")

#Read Excel Inputs
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
    change_dict[value[0] + 0.1] = (str(value[2]) + '_Other Free Text Response')

#Rename columns
pvt_final = pvt_sort.rename(columns=change_dict, inplace=False)
#Write to excel output
pvt_final.to_excel(output_path + '\{}_pivot_{}.xlsx'.format(survey, dt_string))
print("Pivot Complete")




