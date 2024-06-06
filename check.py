from pptx import Presentation
import pandas as pd
import os
import re
from index import summarize_slides

# 月の表紙スライドチェック

import pandas as pd

def check_order(df):
    # Create an empty DataFrame to store the incorrect rows
    incorrect_rows = pd.DataFrame(columns=df.columns)

    # Separate category 2 and 3
    category_2 = df[df['category_number'] == 2]
    category_3 = df[df['category_number'] == 3]

    # Check for order within each category
    category_2_sorted = category_2.sort_values(by=['year', 'month', 'day'])
    category_3_sorted = category_3.sort_values(by=['year', 'month', 'day'])

    # Identify rows that are out of order in category 2
    if not category_2_sorted.equals(category_2):
        incorrect_rows = pd.concat([incorrect_rows, category_2[category_2.reset_index().index != category_2_sorted.reset_index().index]])

    # Identify rows that are out of order in category 3
    if not category_3_sorted.equals(category_3):
        incorrect_rows = pd.concat([incorrect_rows, category_3[category_3.reset_index().index != category_3_sorted.reset_index().index]])

    # Check for consistency between category 2 and 3
    for i, row in category_2_sorted.iterrows():
        next_month = row['month'] + 1
        incorrect_next_month = category_3_sorted[(category_3_sorted['year'] == row['year']) & 
                                                 (category_3_sorted['month'] == next_month) & 
                                                 (category_3_sorted['day'] < row['day'])]
        incorrect_rows = pd.concat([incorrect_rows, incorrect_next_month])
    
    print(incorrect_rows)
    return incorrect_rows


file_path1 = r"【事例資料】LOUIS VUITTON_LINE 公式アカウント_メッセージ配信_2024年1月以降.pptx"
file_path2 = r"【事例資料】ヴァレンティノ_LINE 公式アカウント_メッセージ配信事例_2024年1月以降.pptx"
file_path3 = r"【事例資料】ベイクルーズ_LINE 公式アカウント_メッセージ配信_2024年1月以降.pptx"
# ファイルパスを指定して関数を呼び出し、結果を表示します。
# print(extract_text_from_pptx_by_slide(file_path1))
summarize_slides(file_path1).to_csv('file1Test1.csv')
check_order(summarize_slides(file_path1))
