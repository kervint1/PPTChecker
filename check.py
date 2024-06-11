from pptx import Presentation
import pandas as pd
import os
import re
from index import summarize_slides

# 月の表紙スライドチェック

import pandas as pd

#日付の順番をチェック(未使用)
def check_order(df):
    # category_number が 2 と 3 の行をフィルタリング
    df_filtered = df[df['category_number'].isin([2, 3])]
    
    # 間違っている行を収集するための空のデータフレームを初期化
    incorrect_df = pd.DataFrame(columns=df.columns)

    # ユニークな category_number ごとにイテレート
    for category in df_filtered['category_number'].unique():
        # 特定の category_number に対するデータフレームのサブセットを取得
        category_df = df_filtered[df_filtered['category_number'] == category]
        
        # year, month, day の順にサブセットをソート
        category_df_sorted = category_df.sort_values(by=['year', 'month', 'day'])
        
        # ソートされたデータフレームとソート前のデータフレームを比較して、間違っている行を特定
        if not category_df_sorted.index.equals(category_df.index):
            print(category_df)
            print(category_df_sorted)
            incorrect_df = pd.concat([incorrect_df, category_df[category_df.index != category_df_sorted.index]])

    return incorrect_df

#正しい日付の順序にするための動かすべき行を出力->({"categorynum":(スライドナンバー)-1})
def find_rows_to_move(df):
    # category_number が 2 と 3 の行をフィルタリング
    df_filtered = df[df['category_number'].isin([2, 3])]
    
    # 各カテゴリーの移動すべき行を格納する辞書
    rows_to_move = {}

    # ユニークな category_number ごとにイテレート
    for category in df_filtered['category_number'].unique():
        # 特定の category_number に対するデータフレームのサブセットを取得
        category_df = df_filtered[df_filtered['category_number'] == category]
        
        # year, month, day の順にサブセットをソート
        category_df_sorted = category_df.sort_values(by=['year', 'month', 'day'])

        # ソートされたデータフレームとソート前のデータフレームを比較して、間違っている行を特定
        if not category_df_sorted.index.equals(category_df.index):
            original_index = category_df.index.tolist()
            sorted_index = category_df_sorted.index.tolist()

            # リストをソートするために必要な最小の移動を見つける
            moves = []
            sorted_positions = {v: i for i, v in enumerate(sorted_index)}

            for i in range(len(original_index)):
                while original_index[i] != sorted_index[i]:
                    swap_idx = sorted_positions[original_index[i]]
                    moves.append(original_index[i])
                    original_index[i], original_index[swap_idx] = original_index[swap_idx], original_index[i]

            rows_to_move[category] = moves

    return rows_to_move

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
print(find_single_month_rows(summarize_slides(file_path1)))