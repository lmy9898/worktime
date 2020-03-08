import pandas as pd
from datetime import timedelta,datetime

obj = "C:/Users/test/Desktop/나무발발이/이민영2월.xlsx"
dst ="C:/Users/test/Desktop/나무발발이/이민영.csv"

df = pd.read_excel(obj)
df = df[['발생시각','상태']]

# 새벽 야근을 고려한 6시간 빼기
df.발생시각 = pd.to_datetime(df.발생시각)
delta = timedelta(hours=6)
df['발생시각'] = df.발생시각-delta

# 출입 = 출근
df['상태'] = df.상태.str.split('(').str[0].str[0]

for i in range(df.shape[0]):
  df.loc[i,'발생시각'] = str(df.loc[i,'발생시각'])
# 날짜별 그룹화를 위한 날짜, 시각 분할
df['날짜'], df['시각'] = df.발생시각.str.split(' ').str

df = df.groupby(['날짜','상태']).agg(['max','min'])

df = df.unstack()
df_출 = df[('시각','min','출')]
df_퇴 = df[('시각','max','퇴')]
df_발_퇴 = df[('발생시각','max','퇴')]
df_발_출 = df[('발생시각','min','출')]
df = pd.concat([df_발_출,df_발_퇴,df_출,df_퇴,],axis=1)

df.columns = df.columns.droplevel()
df.columns = df.columns.droplevel()
df.columns = ["발생시각_출", "발생시각_퇴",'출','퇴']

df.reset_index(level=['날짜'], inplace = True)
df.퇴 = df.퇴.fillna('12:00:00')

#근무시간 추가
df.퇴 = pd.to_datetime(df.날짜+' '+df.퇴)
df.출 = pd.to_datetime(df.날짜+' '+df.출)
df['퇴'] = df.퇴+delta
df['출'] = df.출+delta
df['근무시간'] =  df['퇴'] - df['출']
df = df[['날짜','출','퇴','근무시간']]

runtime = str(df.근무시간.mean())[7:15]
for i in range(len(df)):
  df.loc[i,'출'] = str(df.loc[i,'출'])
  df.loc[i,'퇴'] = str(df.loc[i,'퇴'])
  df.loc[i,'근무시간'] = str(df.loc[i,'근무시간'])[7:15]
df['퇴근시각'], df['출근시각'],df['근무시간']= df.퇴.str.split(' ').str[1], df.출.str.split(' ').str[1], df.근무시간.str.split(' ').str[0]

df['야근']=''
df['지각']=''

for i in range(len(df)):

  t1 = int(df.퇴근시각[i][:2])
  if t1 >= 20 or t1< 9:
    df.야근[i] = 'O'

  t2 = float(df.출근시각[i][:5].replace(':','.'))
  if t2 >= 9.3 :
    df.지각[i] = 'O'

df = df[['날짜','출근시각','퇴근시각','근무시간','지각','야근']]

days = len(df)
over_num = df.야근.value_counts()['O']
late_num = df.지각.value_counts()['O']

print('평균 근무 시간 : ',runtime)
print('근무 일 수 : ', days)
print('야근 횟수 : ', over_num )
print('지각 횟수 : ', late_num )

my_dict = {"object": ['평균 근무 시간', '근무 일 수','야근 횟수','지각 횟수'], "value": [runtime,days,over_num,late_num]}
total = pd.DataFrame(my_dict)

a = pd.concat([df,total], axis = 1,ignore_index=False)
a
a.to_csv(dst, mode='w', encoding='euckr')