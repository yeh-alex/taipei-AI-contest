import pandas as pd
from sklearn.model_selection import train_test_split
from sklearn.svm import SVC
from sklearn.preprocessing import StandardScaler
from sklearn.metrics import accuracy_score, classification_report

df1 = pd.read_excel("2020~2021_label.xlsx")
df2 = pd.read_excel("2021~2022_label.xlsx")
df3 = pd.read_excel ("2022~2023_label.xlsx")


df_combined = pd.concat([df1,df2,df3], ignore_index=True)
merged_df = pd.merge(df1, df2,df3 ,how='left', on='會員編號')
# print(df_combined.head())
df_combined.to_excel('2020~2023.xlsx', index=False)

# output_file_name = "2020~2023.xlsx"
# result.to_excel(output_file_name, index=False)

# 讓使用者下載結果檔案
# files.download(output_file_name)