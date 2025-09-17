##### 2020~2021
import pandas as pd

a = '2020'
b = '2021'

# 读取 a.xlsx 和 b.xlsx 两个文件
a_df = pd.read_excel(a + '.xlsx')
b_df = pd.read_excel(b + '.xlsx')

# 整理資料的函數
def process_data(df, year):
    # 清理並拆解 '半年區間' 以獲取日期範圍的開始日期
    df[['半年區間開始', '半年區間結束']] = df['半年區間'].str.split('~', expand=True)
    
    # 將 '半年區間開始' 和 '半年區間結束' 轉換為日期格式
    df['半年區間開始'] = pd.to_datetime(df['半年區間開始'], format='%Y/%m/%d', errors='coerce')
    df['半年區間結束'] = pd.to_datetime(df['半年區間結束'], format='%Y/%m/%d', errors='coerce')
    
    # 移除無效的 '半年區間開始' 和 '半年區間結束' 行
    df = df.dropna(subset=['半年區間開始', '半年區間結束'])
    
    # 用lambda根據開始日期的月份來分類為上半年或下半年區間
    df['半年區間分類'] = df['半年區間開始'].apply(
        lambda x: f'{year}/1/1~{year}/6/30' if x.month <= 6 else f'{year}/7/1~{year}/12/31'
    )

    # 按 客戶名稱 和 半年區間分類 分組，並對每個分組的半年區間分類後次數做總和
    processed_df = df.groupby(['客戶名稱', '半年區間分類']).agg(
        次數=('次數', 'sum'),
    ).reset_index()

    # 按 客戶名稱 和 半年區間分類 分組，並對每個分組的半年區間分類後金額做總和
    processed_df2 = df.groupby(['客戶名稱', '半年區間分類']).agg(
        金額=('金額', 'sum')
    ).reset_index()

    # Pivot: 轉換為上半年與下半年兩列，處理次數
    processed_df = processed_df.pivot_table(
        index=['客戶名稱'],
        columns='半年區間分類',
        values='次數',
        aggfunc='sum',
        fill_value=0
    ).reset_index()

    # Pivot: 轉換為上半年與下半年兩列，處理金額
    processed_df2 = processed_df2.pivot_table(
        index=['客戶名稱'],
        columns='半年區間分類',
        values='金額',
        aggfunc='sum',
        fill_value=0
    ).reset_index()

    # 重設列名稱，使其更清晰
    processed_df.columns.name = None
    processed_df2.columns.name = None

    # 重新命名次數列
    processed_df.rename(columns={
        f'{year}/1/1~{year}/6/30': '上半年次數',
        f'{year}/7/1~{year}/12/31': '下半年次數',
    }, inplace=True)

    # 重新命名金額列
    processed_df2.rename(columns={
        f'{year}/1/1~{year}/6/30': '上半年金額',
        f'{year}/7/1~{year}/12/31': '下半年金額'
    }, inplace=True)

    # 合併次數和金額資料
    merged_df = pd.merge(processed_df, processed_df2, on='客戶名稱', how='left')
    
    return merged_df


# 将两个资料表进行整理
processed_df_2021 = process_data(a_df, a)
processed_df_2022 = process_data(b_df, b)

# 確保 '客戶名稱' 在兩個 DataFrame 中都是字串型別
processed_df_2021['客戶名稱'] = processed_df_2021['客戶名稱'].astype(str)
processed_df_2022['客戶名稱'] = processed_df_2022['客戶名稱'].astype(str)

# 左连接 (LEFT JOIN) 两个资料表，依据 '客戶名稱' 进行合并
merged_df = pd.merge(processed_df_2021, processed_df_2022, on='客戶名稱', how='left', suffixes=('_'+a, '_'+b))

# 从 b_df 中获取最新的 '客戶代號'（最后出现的）
df_b_unique = b_df[['客戶名稱', '客戶代號']].drop_duplicates(subset=['客戶名稱'], keep='last')

# 从 a_df 中获取最新的 '客戶代號'（最后出现的）
df_a_unique = a_df[['客戶名稱', '客戶代號']].drop_duplicates(subset=['客戶名稱'], keep='last')

# 合併 '客戶代號' 資料，首先合併 b_df 的最新 '客戶代號'
merged_df = pd.merge(merged_df, df_b_unique[['客戶名稱', '客戶代號']], on='客戶名稱', how='left')

# 对于没有在 b_df 中找到的客戶名稱，从 a_df 中补充 '客戶代號'
merged_df['客戶代號'] = merged_df['客戶代號'].fillna(merged_df['客戶名稱'].map(df_a_unique.set_index('客戶名稱')['客戶代號']))

# 将 2021 和 2022 的上下半年次数字段，如果值不是0，设为1
merged_df[a+'年總次數'] = merged_df['上半年次數_'+a] + merged_df['下半年次數_'+a]
merged_df[a+'年總金額'] = merged_df['上半年金額_'+a] + merged_df['下半年金額_'+a]

merged_df[f'上半年次數_{a}'] = merged_df[f'上半年次數_{a}'].apply(lambda x: 1 if x > 0 else 0)
merged_df[f'下半年次數_{a}'] = merged_df[f'下半年次數_{a}'].apply(lambda x: 1 if x > 0 else 0)
merged_df[f'上半年次數_{b}'] = merged_df[f'上半年次數_{b}'].apply(lambda x: 1 if x > 0 else 0)
merged_df[f'下半年次數_{b}'] = merged_df[f'下半年次數_{b}'].apply(lambda x: 1 if x > 0 else 0)

# 新增一列，标记 2022 上半年或下半年是否有值（布尔值）
merged_df['布林值'] = merged_df[f'上半年次數_{b}'] + merged_df[f'下半年次數_{b}']

# 最终整理，只保留需要的
final_result = merged_df[['客戶代號', '客戶名稱', '上半年次數_'+a, '下半年次數_'+a, '上半年次數_'+b, '下半年次數_'+b, '布林值', a+'年總次數', a+'年總金額']]

# 顯示結果
#print(final_result)

# 如果需要将结果存成 Excel 可以这么做
final_result.to_excel(a+'~'+b+'_label.xlsx', index=False)


##### 2021~2024
'''
import pandas as pd

a = '2020'
b = '2021'

# 读取 a.xlsx 和 b.xlsx 两个文件
a_df = pd.read_excel(a + '.xlsx')
b_df = pd.read_excel(b + '.xlsx')

# 整理資料的函數
def process_data(df, year):
    # 清理並拆解 '半年區間' 以獲取日期範圍的開始日期
    df[['半年區間開始', '半年區間結束']] = df['半年區間'].str.split('~', expand=True)
    
    # 將 '半年區間開始' 和 '半年區間結束' 轉換為日期格式
    df['半年區間開始'] = pd.to_datetime(df['半年區間開始'], format='%Y/%m/%d', errors='coerce')
    df['半年區間結束'] = pd.to_datetime(df['半年區間結束'], format='%Y/%m/%d', errors='coerce')
    
    # 移除無效的 '半年區間開始' 和 '半年區間結束' 行
    df = df.dropna(subset=['半年區間開始', '半年區間結束'])
    
    # 用lambda根據開始日期的月份來分類為上半年或下半年區間
    df['半年區間分類'] = df['半年區間開始'].apply(
        lambda x: f'{year}/1/1~{year}/6/30' if x.month <= 6 else f'{year}/7/1~{year}/12/31'
    )

    # 按 客戶代號 和 客戶名稱 分組，並對每個分組的半年區間分類後次數做總和
    processed_df = df.groupby(['客戶代號', '客戶名稱', '半年區間分類']).agg(
        次數=('次數', 'sum'),
    ).reset_index()

    # 按 客戶代號 和 客戶名稱 分組，並對每個分組的半年區間分類後金額做總和
    processed_df2 = df.groupby(['客戶代號', '客戶名稱', '半年區間分類']).agg(
        金額=('金額', 'sum')
    ).reset_index()
    
    # 定義所有可能的 '客戶代號', '客戶名稱' 和 '半年區間分類' 組合
    all_combinations = pd.MultiIndex.from_product(
        [df['客戶代號'].unique(), df['客戶名稱'].unique(), df['半年區間分類'].unique()],
        names=['客戶代號', '客戶名稱', '半年區間分類']
    )
    
    # Pivot: 轉換為上半年與下半年兩列，處理次數
    processed_df = processed_df.pivot_table(
        index=['客戶代號', '客戶名稱'],
        columns='半年區間分類',
        values='次數',
        aggfunc='sum',
        fill_value=0
    ).reset_index()

    # Pivot: 轉換為上半年與下半年兩列，處理金額
    processed_df2 = processed_df2.pivot_table(
        index=['客戶代號', '客戶名稱'],
        columns='半年區間分類',
        values='金額',
        aggfunc='sum',
        fill_value=0
    ).reset_index()
    
    # 重設列名稱，使其更清晰
    processed_df.columns.name = None
    processed_df2.columns.name = None

    # 重新命名次數列
    processed_df.rename(columns={
        f'{year}/1/1~{year}/6/30': '上半年次數',
        f'{year}/7/1~{year}/12/31': '下半年次數',
    }, inplace=True)

    # 重新命名金額列
    processed_df2.rename(columns={
        f'{year}/1/1~{year}/6/30': '上半年金額',
        f'{year}/7/1~{year}/12/31': '下半年金額'
    }, inplace=True)

    # 合併次數和金額資料
    merged_df = pd.merge(processed_df, processed_df2, on=['客戶代號', '客戶名稱'], how='left')
    
    return merged_df



# 将两个资料表进行整理
processed_df_2021 = process_data(a_df, a)
processed_df_2022 = process_data(b_df, b)


#print(processed_df_2022)
# 確保 '客戶代號' 欄位在兩個 DataFrame 中都是字串型別
processed_df_2021['客戶代號'] = processed_df_2021['客戶代號'].astype(str)
processed_df_2022['客戶代號'] = processed_df_2022['客戶代號'].astype(str)
# 左连接 (LEFT JOIN) 两个资料表，依据 '客戶代號' 和 '客戶名稱' 进行合并
merged_df = pd.merge(processed_df_2021, processed_df_2022, on=['客戶代號', '客戶名稱'], how='left', suffixes=('_'+a, '_'+b))

# 将 2021 和 2022 的上下半年次数字段，如果值不是0，设为1
merged_df[a+'年總次數'] = merged_df['上半年次數_'+a] + merged_df['下半年次數_'+a]
merged_df[a+'年總金額'] = merged_df['上半年金額_'+a] + merged_df['下半年金額_'+a]

merged_df[f'上半年次數_{a}'] = merged_df[f'上半年次數_{a}'].apply(lambda x: 1 if x > 0 else 0)
merged_df[f'下半年次數_{a}'] = merged_df[f'下半年次數_{a}'].apply(lambda x: 1 if x > 0 else 0)
merged_df[f'上半年次數_{b}'] = merged_df[f'上半年次數_{b}'].apply(lambda x: 1 if x > 0 else 0)
merged_df[f'下半年次數_{b}'] = merged_df[f'下半年次數_{b}'].apply(lambda x: 1 if x > 0 else 0)

# 新增一列，标记 2022 上半年或下半年是否有值（布尔值）
#merged_df['布林值'] = (merged_df['上半年次數_2024'] | merged_df['下半年次數_2024']).astype(int)
#merged_df['布林值'] = merged_df['上半年次數_'+b] + merged_df['下半年次數_'+b]
merged_df['布林值'] = merged_df[f'上半年次數_{b}'] + merged_df[f'下半年次數_{b}']

# 最终整理，只保留需要的
final_result = merged_df[['客戶代號', '客戶名稱', '上半年次數_'+a, '下半年次數_'+a, '上半年次數_'+b, '下半年次數_'+b, '布林值', a+'年總次數', a+'年總金額']]

# 顯示結果
#print(final_result)

# 如果需要将结果存成 Excel 可以这么做
final_result.to_excel(a+'~'+b+'_label.xlsx', index=False)
'''
