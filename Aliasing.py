import pandas as pd

df_ref = pd.read_excel('C:\Python\TestProject\QuestionRef.xlsx', 0)
df_main = pd.read_excel('C:\Python\TestProject\About_Me_Test_Output_pivot_202101131309.xlsx')

df_ref_cln = df_ref.dropna(axis=0, how='all')
#df_alias = df_ref_cln['Alias']
print(df_main)
print(df_ref_cln)