from pptx import Presentation
import pandas as pd
from index import summarize_latest_slides

# 月の表紙スライドチェック

import pandas as pd

def find_misplaced_row_index(df):
    # category_number が 2 または 3 の行をフィルタリング
    category_2_3_df = df[df['category_number'].isin([2, 3])]

    # date 列でソート
    category_2_3_df_sorted = category_2_3_df.sort_values(by='date')

    # 元のインデックスとソートされたインデックスを比較
    original_index = category_2_3_df.index
    sorted_index = category_2_3_df_sorted.index

    # インデックスが異なる行を特定
    mismatched_indices = original_index[original_index != sorted_index]

    # 異なるインデックスが存在する場合、最初のものを返す
    if not mismatched_indices.empty:
        return mismatched_indices[0]  # 行のインデックスを返す
    else:
        return None

#表紙チェック
def month_check(df):
    # category_number が 2 の行の月を取得
    category_2_months = df[df['category_number'] == 2]['date'].dt.month.unique()

    # category_number が 3 の行の月を取得
    category_3_months = df[df['category_number'] == 3]['date'].dt.month.unique()

    # category_number 2 にあって category_number 3 にない月
    category_2_not_in_3 = set(category_2_months) - set(category_3_months)

    # category_number 3 にあって category_number 2 にない月
    category_3_not_in_2 = set(category_3_months) - set(category_2_months)

    return {
        "category_2_not_in_3": list(category_2_not_in_3),
        "category_3_not_in_2": list(category_3_not_in_2)
    }

def arrow_check(df):
    messages = []
    category_3_df = df[df['category_number'] == 3]
    
    for idx, row in category_3_df.iterrows():
        if pd.isna(row['arrow_presence']):
            if row['ad_number_count'] != 0 or row['lp_count'] != 0 or row['lp_number_count'] != 0:
                messages.append(f"{idx + 1}スライドで矢印にミスがあります。")
        else:
            if row['ad_number_count'] == 0 or row['lp_count'] == 0 or row['lp_number_count'] == 0:
                messages.append(f"{idx + 1}スライドで矢印にミスがあります。")
    
    return messages

def number_check(df):
    messages = []
    category_3_df = df[df['category_number'] == 3]
    
    for idx, row in category_3_df.iterrows():
        if row['ad_number_count'] != row['lp_number_count']:
            messages.append(f"{idx + 1}スライドで振り番にミスがあります")
        if row['lp_number_count'] < row['lp_count']:
            messages.append(f"{idx + 1}スライドで振り番にミスがあります")
    
    return messages

def duplicates_check(df):
    messages = []
    category_3_df = df[df['category_number'] == 3]
    duplicates = category_3_df[category_3_df.duplicated(subset='date', keep=False)]
    
    for idx in duplicates.index:
        messages.append(f"{idx + 1}スライドで同じ日時のスライドがあります")
    
    return messages

def account_check(df):
    messages = []
    category_3_df = df[df['category_number'] == 3]
    first_account_name = category_3_df.iloc[0]['account_name']
    
    for idx, row in category_3_df.iterrows():
        if row['account_name'] != first_account_name:
            messages.append(f"{idx + 1}スライドでアカウント名が違います")
    
    return messages

def message_voom_check(df):
    messages = []
    first_message_or_voom = df.iloc[0]['message_or_voom']
    
    for idx, row in df.iterrows():
        if row['message_or_voom'] != first_message_or_voom:
            messages.append(f"{idx + 1}スライドでメッセージorVOOMが違います")
    
    return messages

def ad_check(df):
    messages = []
    category_3_df = df[df['category_number'] == 3]
    
    for idx, row in category_3_df.iterrows():
        if not row['ad_presence']:
            messages.append(f"{idx + 1}スライドでadがありません")
    
    return messages



#全チェック
def all_checks(df):
    messages = []

    misplaced_index = find_misplaced_row_index(df)
    if misplaced_index is not None:
        messages.append(f"スライド{misplaced_index + 1} から順番が違います")

    month_check_result = month_check(df)
    if month_check_result["category_2_not_in_3"] or month_check_result["category_3_not_in_2"]:
        messages.append(f"{list(month_check_result['category_2_not_in_3'])}月分のスライドがありません")
        messages.append(f"{list(month_check_result['category_3_not_in_2'])}月分のスライドがありません")

    messages.extend(arrow_check(df))
    messages.extend(number_check(df))
    messages.extend(duplicates_check(df))
    messages.extend(account_check(df))
    messages.extend(message_voom_check(df))
    messages.extend(ad_check(df))

    return messages



# file_path1 = r"【事例資料】LOUIS VUITTON_LINE 公式アカウント_メッセージ配信_2024年1月以降.pptx"
# file_path2 = r"【事例資料】ヴァレンティノ_LINE 公式アカウント_メッセージ配信事例_2024年1月以降.pptx"
# file_path3 = r"【事例資料】ベイクルーズ_LINE 公式アカウント_メッセージ配信_2024年1月以降.pptx"
# ファイルパスを指定して関数を呼び出し、結果を表示します。
# print(extract_text_from_pptx_by_slide(file_path1))
# summarize_slides(file_path1).to_csv('file1Test1.csv')
# print(find_single_month_rows(summarize_slides(file_path1)))


# df = summarize_latest_slides(file_path1,False)