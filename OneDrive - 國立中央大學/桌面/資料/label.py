import  pandas as pd
#import  
data = pd.read_excel("電行數據-通話紀錄20200101-20231231.xlsx")

print(data.head())

data[['通話狀態1']] = data[['通話狀態1']].map({'推薦舊產品': 1, '關心聯絡': 0, '已接通': 1, '無法接通': 0, '無法撥打': 0})
print(data.head())
