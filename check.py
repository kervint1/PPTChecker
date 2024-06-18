from pptx import Presentation
import pandas as pd
from index import summarize_slides

# 月の表紙スライドチェック

import pandas as pd

def find_misplaced_row_index(df):
    def compare_and_find_mismatch(original_df, sorted_df):
        original_index = original_df.index.tolist()
        sorted_index = sorted_df.index.tolist()
        
        for i, (orig_idx, sorted_idx) in enumerate(zip(original_index, sorted_index)):
            if orig_idx != sorted_idx:
                return orig_idx
        return None

    # category_number が 2 または 3 の行をフィルタリング
    category_2_3_df = df[df['category_number'].isin([2, 3])]

    # date 列でソート
    category_2_3_df_sorted = category_2_3_df.sort_values(by='date')

    # ソート前後の DataFrame を比較して、最初に異なる行を特定
    mismatch_index = compare_and_find_mismatch(category_2_3_df, category_2_3_df_sorted)

    if mismatch_index is not None:
        return category_2_3_df.loc[mismatch_index]
    else:
        return None

#表紙チェック
def find_single_month_rows(df):
    # Group by 'month' and count occurrences
    month_counts = df['month'].value_counts()
    
    # Identify months with only one occurrence
    single_months = month_counts[month_counts == 1].index.tolist()
    
    # Filter the dataframe to get rows with single-occurrence months
    single_month_rows = df[df['month'].isin(single_months)]
    
    return single_month_rows

#全チェック()
def all_checks(df):
    df = summarize_slides(df)
    # 移動すべき行を見つける
    rows_to_move = find_rows_to_move(df)
    
    # 月が1つしかない行を見つける
    single_month_rows = find_single_month_rows(df)
    
    # 移動すべき行をデータフレームに変換
    move_rows_df = pd.DataFrame([(key, row) for key, rows in rows_to_move.items() for row in rows], columns=['category_number', 'row_index'])
    
    # 月が1つしかない行と結合
    combined_df = pd.concat([single_month_rows, df.loc[move_rows_df['row_index']]]).drop_duplicates().reset_index(drop=True)
    
    return combined_df

file_path1 = r"【事例資料】LOUIS VUITTON_LINE 公式アカウント_メッセージ配信_2024年1月以降.pptx"
file_path2 = r"【事例資料】ヴァレンティノ_LINE 公式アカウント_メッセージ配信事例_2024年1月以降.pptx"
file_path3 = r"【事例資料】ベイクルーズ_LINE 公式アカウント_メッセージ配信_2024年1月以降.pptx"
# ファイルパスを指定して関数を呼び出し、結果を表示します。
# print(extract_text_from_pptx_by_slide(file_path1))
# summarize_slides(file_path1).to_csv('file1Test1.csv')
# print(find_single_month_rows(summarize_slides(file_path1)))


df = summarize_slides(file_path1)
print(find_misplaced_row_index(df))