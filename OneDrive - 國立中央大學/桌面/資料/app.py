'''
import pandas as pd
import os

# 取得當前程式所在資料夾的路徑
current_dir = os.path.dirname(os.path.realpath(__file__))

# 指定 Excel 檔案名稱
file_name = "復亨銷售紀錄-2021.xlsx"

# 完整路徑
file_path = os.path.join(current_dir, file_name)

# 讀取 Excel 檔案
if os.path.exists(file_path):
    df = pd.read_excel(file_path)
    
    # 確保銷貨日期欄位是日期型別
    df['銷貨日期'] = pd.to_datetime(df['銷貨日期'])

    # 新增一個欄位，標註是上半年還是下半年，根據年份動態決定
    df['半年區間'] = df['銷貨日期'].apply(
        lambda x: f"{x.year}/1/1~{x.year}/6/30" if x.month <= 6 else f"{x.year}/7/1~{x.year}/12/31"
    )
    
    # 以客戶代號、半年區間進行分組並統計每個區間的次數
    #result = df.groupby(['半年區間', '客戶代號', '客戶名稱']).size().reset_index(name='次數')
    
    result = df.groupby(['半年區間', '客戶代號', '客戶名稱']).agg(
        次數=('銷貨單號', 'size'),
        金額=('發票金額', 'sum')
    ).reset_index()
    
    
    # 輸出結果為新的 Excel 檔案
    output_file_path = os.path.join(current_dir, "2021.xlsx")
    result.to_excel(output_file_path, index=False)

    print(f"結果已成功輸出到: {output_file_path}")
    
else:
    print(f"找不到檔案: {file_name}")
'''

import pandas as pd
import os

# 取得當前程式所在資料夾的路徑
current_dir = os.path.dirname(os.path.realpath(__file__))

# 指定 Excel 檔案名稱
file_name = "復亨銷售紀錄-2021.xlsx"

# 完整路徑
file_path = os.path.join(current_dir, file_name)

# 讀取 Excel 檔案
if os.path.exists(file_path):
    df = pd.read_excel(file_path)
    
    # 確保銷貨日期欄位是日期型別
    df['銷貨日期'] = pd.to_datetime(df['銷貨日期'])

    # 新增一個欄位，標註是上半年還是下半年，根據年份動態決定
    df['半年區間'] = df['銷貨日期'].apply(
        lambda x: f"{x.year}/1/1~{x.year}/6/30" if x.month <= 6 else f"{x.year}/7/1~{x.year}/12/31"
    )

    # 先將客戶代號轉為字串，並去除可能的空白字符
    df['客戶代號'] = df['客戶代號'].astype(str).str.strip()
    df['客戶名稱'] = df['客戶名稱'].str.strip()  # 確保客戶名稱不受多餘空白字符影響
    
    # 以半年區間和客戶名稱進行分組，並對於相同的名稱進行加總
    result = df.groupby(['半年區間', '客戶名稱']).agg(
        客戶代號=('客戶代號', 'first'),  # 取每個分組中第一個客戶代號
        次數=('銷貨單號', 'size'),        # 次數為銷貨單號的數量
        金額=('發票金額', 'sum')          # 金額為發票金額的總和
    ).reset_index()

    # 輸出結果為新的 Excel 檔案
    output_file_path = os.path.join(current_dir, "2021.xlsx")
    result.to_excel(output_file_path, index=False)

    print(f"結果已成功輸出到: {output_file_path}")
    
else:
    print(f"找不到檔案: {file_name}")


