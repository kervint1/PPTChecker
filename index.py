from pptx import Presentation
import pandas as pd
import os

def extract_text_from_pptx_by_slide(file_path):
    slides_texts = []
    prs = Presentation(file_path)
    for slide in prs.slides:
        slide_texts = []  # スライドごとのテキストを格納するリスト
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                slide_texts.append(shape.text)
        slides_texts.append(slide_texts)
    return slides_texts

def classify_slide(slide):
    """
    スライドを分類する関数。表紙、月のスライド、内容のスライドのいずれかに分類。
    """
    objects = slide.shapes
    for shape in objects:
        if shape.top <= 100 and shape.bottom >= 200:
            return 'cover'
        elif shape.top <= 200 and shape.bottom >= 200:
            return 'month'
        elif shape.top <= 100 and shape.bottom >= 300:
            return 'content'
    return None

def extract_cover_data(slide):
    """
    表紙のスライドからデータを抽出する関数。
    """
    # 実際の実装はここに
    account_name = "example_account"
    error_message = "no errors"
    return pd.DataFrame([{
        'category_number': 1,
        'account_name': account_name,
        'error_message': error_message
    }])

def extract_month_data(slide):
    """
    月のスライドからデータを抽出する関数。
    """
    # 実際の実装はここに
    account_name = "example_account"
    year = 2023
    month = 5
    error_message = "no errors"
    return pd.DataFrame([{
        'category_number': 2,
        'account_name': account_name,
        'year': year,
        'month': month,
        'error_message': error_message
    }])

def extract_content_data(slide):
    """
    内容のスライドからデータを抽出する関数。
    """
    # 実際の実装はここに
    account_name = "example_account"
    message_or_voom = "message"
    month = 5
    day = 15
    datetime = "2023-05-15 10:00"
    ad_presence = True
    ad_account_name = "ad_account"
    lp_count = 3
    lp_number_count = 5
    arrow_presence = False
    error_message = "no errors"
    return pd.DataFrame([{
        'category_number': 3,
        'account_name': account_name,
        'message_or_voom': message_or_voom,
        'month': month,
        'day': day,
        'datetime': datetime,
        'ad_presence': ad_presence,
        'ad_account_name': ad_account_name,
        'lp_count': lp_count,
        'lp_number_count': lp_number_count,
        'arrow_presence': arrow_presence,
        'error_message': error_message
    }])

def summarize_slides(slides):
    """
    スライドを分類し、それぞれのデータを取得してpandasでデータフレームにまとめる関数。
    """
    data_frames = []

    for slide in slides:
        slide_type = classify_slide(slide)
        if slide_type == 'cover':
            df = extract_cover_data(slide)
        elif slide_type == 'month':
            df = extract_month_data(slide)
        elif slide_type == 'content':
            df = extract_content_data(slide)
        else:
            df = pd.DataFrame([{
                'category_number': 4,
                'account_name': None,
                'error_message': 'No slide content'
            }])
        data_frames.append(df)

    result_df = pd.concat(data_frames, ignore_index=True)
    return result_df
    

def save_texts_to_excel(texts, filename):
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        for i, (file_name, file_texts) in enumerate(texts):
            all_texts = []
            for slide_index, slide_texts in enumerate(file_texts):
                for text in slide_texts:
                    all_texts.append({'Slide Number': slide_index + 1, 'Text': text})
            df = pd.DataFrame(all_texts)
            if not df.empty:
                df.to_excel(writer, sheet_name=f'{os.path.splitext(file_name)[0]}', index=False)


file_path1 = r"【事例資料】LOUIS VUITTON_LINE 公式アカウント_メッセージ配信_2024年1月以降.pptx"
# ファイルパスを指定して関数を呼び出し、結果を表示します。
print(extract_text_from_pptx_by_slide(file_path1))
