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

def classify_slide(slide,top_list,left_list):
    """
    スライドを分類する関数。表紙、月のスライド、内容のスライドのいずれかに分類。
    """
    objects = slide.shapes
    permissible = 5

    for shape in objects:
        if abs(shape.top.pt-top_list[0]) < permissible and abs(shape.left.pt-left_list[0]) < permissible:
            return 'cover'
        elif abs(shape.top.pt-top_list[1]) < permissible and abs(shape.left.pt-left_list[1]) < permissible:
            return 'month'
        elif abs(shape.top.pt-top_list[2]) < permissible and abs(shape.left.pt-left_list[2]) < permissible:
            return 'content'
    return None

def extract_cover_data(slide):
    """
    表紙のスライドからデータを抽出する関数。
    """
    objects = slide.shapes
    permissible = 5
    cover_position_top = []
    cover_position_left = []

    for shape in objects:
        if abs(shape.top.pt-cover_position_top[0]) < permissible and abs(shape.left.pt-cover_position_left[0]) < permissible:
            account_name = shape.text
        elif abs(shape.top.pt-cover_position_top[1]) < permissible and abs(shape.left.pt-cover_position_left[1]) < permissible:
            error_message = "no errors"
    

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

def summarize_slides(file_path):
    """
    スライドを分類し、それぞれのデータを取得してpandasでデータフレームにまとめる関数。
    """
    slides = Presentation(file_path).slides

    #【事例資料】LOUIS VUITTON_LINE 公式アカウント_メッセージ配信_2024年1月以降.pptx(少数点以下切り捨て)
    # 上記のpptでpositionを大まかに決める
    # 1 4,2 6,3 6
    LV_position_top = [474,233,3]
    LV_position_left = [311,109,21]
    permissible = 20
    standard_top = []
    standard_left = []
    # その初期３スライドのpptの基準を決める
    for i in range(0,3):
        for shape in slides[i].shapes:
            if (abs(shape.left.pt-LV_position_left[i])<permissible and
                abs(shape.top.pt-LV_position_top[i])<permissible):
                standard_top.append(round(shape.top.pt,0))
                standard_left.append(round(shape.left.pt))
                break
            else:
                None
    if len(standard_top)!=3:
        print("top3 Slide Error")

    print(standard_top,standard_left)
    data_frames = []

    for slide in slides:
        slide_type = classify_slide(slide,standard_top,standard_left)
        if slide_type == 'cover':
            # df = extract_cover_data(slide)
            print(0)
        elif slide_type == 'month':
            # df = extract_month_data(slide)
            print(1)
        elif slide_type == 'content':
            # df = extract_content_data(slide)
            print(2)
        else:
            print(3)
            # df = pd.DataFrame([{
            #     'category_number': 4,
            #     'account_name': None,
            #     'error_message': 'No slide content'
            # }])
        # data_frames.append(df)
        
    
    return 0

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
file_path2 = r"【事例資料】ヴァレンティノ_LINE 公式アカウント_メッセージ配信事例_2024年1月以降.pptx"
file_path3 = r"【事例資料】ベイクルーズ_LINE 公式アカウント_メッセージ配信_2024年1月以降.pptx"
# ファイルパスを指定して関数を呼び出し、結果を表示します。
# print(extract_text_from_pptx_by_slide(file_path1))
summarize_slides(file_path3)

