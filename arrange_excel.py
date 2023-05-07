import os
import datetime as dt
import subprocess
import shutil
import pandas as pd
# print('input date(ex. 230507) of master_db')
# datestr=input('>>')
datestr='230507'
file_org='data_db//master_db'+datestr+'.xlsx'
file_exl='master_list'+datestr+'.xlsx'
file_db='list'+datestr+'.xlsx'
file_map='map'+datestr+'.xlsx'
file_map2='map_r'+datestr+'.xlsx'
df_master=pd.read_excel(file_org, sheet_name='db', header=2)
df_master.to_excel(file_exl)
# subprocess.Popen(["start", "", file_exl], shell=True)
df_extract=df_master.query('履歴DB掲載.str.contains("〇")', engine='python')
df_list=df_extract[['記号', 'NO.', '組織', '年月', '改正種別', '規格番号', '規格タイトル', '改正内容', '参照規格', 'マップ表示(清書)', '規格種別']]
df_list=df_list.sort_values(['年月'], ascending=[True])
# df_list.head()
df_list.reset_index(inplace=True)
df_list['Num']=df_list.reset_index().index+1
df_list=df_list[['Num', '記号', 'NO.', '組織', '年月', '改正種別', '規格番号', '規格タイトル', '改正内容', '参照規格', 'マップ表示(清書)', '規格種別']]
df_list['年月']=pd.to_datetime(df_list['年月'])
df_list['年月2']=df_list['年月'].dt.strftime('%Y年%m月')
df_list=df_list[['Num', '年月', '記号', 'NO.', '組織', '年月2', '改正種別', '規格番号', '規格タイトル', '改正内容', '参照規格', 'マップ表示(清書)', '規格種別']]
df_list.to_excel(file_db)
subprocess.Popen(["start", "", file_db], shell=True)