import pandas as pd
import os
df = pd.DataFrame()
for filename in os.listdir('.'):
    if filename.endswith(".xls"):
        #print(filename)
        xls = pd.ExcelFile(filename)
        #print(xls)
        df_project_code = pd.read_excel(xls,sheet_name = 'Defect Log',skiprows=7,nrows=1,header=None,index_col=None,usecols=[3])
        project_code = df_project_code[0].values[0]
        xs = pd.read_excel(xls,sheet_name = 'Defect Log',skiprows=12 )
        xs.rename(columns={'Unnamed: 1':'Defect_log_no','Unnamed: 2':'Peer_review_Date','Unnamed: 3':'Type'},inplace=True)
        xs = xs.loc[:,'Defect_log_no'::]
        xs['Critical'].fillna(0,inplace=True)
        xs['Major'].fillna(0,inplace=True)
        xs['Minor'].fillna(0,inplace=True)
        xs=xs[xs.Peer_review_Date.notnull()]
        xs['project_code'] = project_code
        #print(xs.columns)
        df = df.append(xs)
df = pd.DataFrame(df).set_index('Defect_log_no')
#df['Peer_review_Date'] = df['Peer_review_Date'].dt.strftime('%Y-%m-%d')
df.profile_report()

df.to_excel("dfl_summary.xlsx")     