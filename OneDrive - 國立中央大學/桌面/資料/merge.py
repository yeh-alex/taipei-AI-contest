import pandas as pd

df1 = pd.read_excel('2020~2021_label.xlsx')  # 第一張表：客戶資料
df2 = pd.read_excel('通話精簡.xlsx')  # 第二張表：通話紀錄

df1['客戶名稱'] = df1['客戶名稱'].str.strip()
df2['客戶名稱'] = df2['客戶名稱'].str.strip()
# df2_grouped = df2.groupby('會員姓名').agg({
#     '通話時間(秒)': 'sum',
#     '通話狀態1': lambda x: x.mode()[0] if not x.mode().empty else None,
#     '通話狀態2': lambda x: x.mode()[0] if not x.mode().empty else None
# }).reset_index()

merged_df = pd.merge(df1, df2, how='left', on='客戶名稱')
print(merged_df.head())