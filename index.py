from pptx import Presentation
import pandas as pd
import os
import re

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

def check_position(shape,permissible,cover_position_top,cover_position_left):
    if (abs(shape.top.pt-cover_position_top) < permissible and 
        abs(shape.left.pt-cover_position_left) < permissible):
        return True
    else :
        return False

def extract_cover_data(slide):
    """
    表紙のスライドからデータを抽出する関数。
    """
    objects = slide.shapes
    permissible = 20
    cover_position_top = [199,233]
    cover_position_left = [191,109]
    account_name = None
    message_or_voom = None
    error_message = None
    for shape in objects:
        if check_position(shape,permissible,cover_position_top[0],cover_position_left[0]):
            account_name = shape.text
        elif check_position(shape,permissible,cover_position_top[1],cover_position_left[1]):
            if re.search(r"メッセージ",shape.text):
                message_or_voom = 1
            elif re.search(r"VOOM",shape.text):
                message_or_voom = 2
        if account_name and message_or_voom:
            break
    if not(account_name and message_or_voom):
        error_message = "オブジェクトが基準値より20pt離れている"
    print(account_name,message_or_voom,error_message)
    # 実際の実装はここに
    # return pd.DataFrame([{
    #     'category_number': 1,
    #     'account_name': account_name,
    #     'message_or_voom': message_or_voom,        
    #     'error_message': error_message
    # }])

def extract_month_data(slide):
    """
    月のスライドからデータを抽出する関数。
    """
    # 実際の実装はここに
    objects = slide.shapes
    permissible = 5
    cover_position_top = [199,233,283]
    cover_position_left = [191,109,111]
    account_name = None
    year = None
    month = None
    error_message = None

    for shape in objects:
        if check_position(shape,permissible,cover_position_top[0],cover_position_left[0]):
            account_name = shape.text
        elif check_position(shape,permissible,cover_position_top[1],cover_position_left[1]):
            if re.search(r"メッセージ",shape.text):
                message_or_voom = 1
            elif re.search(r"VOOM",shape.text):
                message_or_voom = 2
        elif check_position(shape,permissible,cover_position_top[2],cover_position_left[2]):
            match = re.search(r"(\d{4})年(\d{1,2})月", shape.text)
            if match:
                year = int(match.group(1))
                month = int(match.group(2))
    if not(account_name and message_or_voom):
        error_message = "オブジェクトが基準値より20pt離れている"
    print(account_name,message_or_voom,year,month,error_message)

    # return pd.DataFrame([{
    #     'category_number': 2,
    #     'account_name': account_name,
    #     'message_or_voom': message_or_voom,  
    #     'year': year,
    #     'month': month,
    #     'error_message': error_message
    # }])

def extract_content_data(slide):
    """
    内容のスライドからデータを抽出する関数。
    """
    # 実際の実装はここに
    objects = slide.shapes
    normal_permissible = 5
    strict_permissible = 1
    find_permissible =30
    cover_position_top = [145,199]
    cover_position_left = [256,191]
    reference_line_top = [136,145,321,332]
    reference_line_left = [21,35,146,256,356,455,554,653]
    ad2_bottom = None
    
    account_name = None
    message_or_voom = None
    month = None
    day = None
    time = None
    ad_presence = None
    ad_account_name = "ad_account"
    ad_number_count = 0
    lp_count = 0
    lp_number_count = 0
    arrow_presence = None
    error_message = ""
        

    for shape in objects:
        if check_position(shape,normal_permissible,cover_position_top[0],cover_position_left[0]):  
            pattern = r"(.+?)\s+LINE公式アカウント\s+(.+?)\s+活用状況"
            match = re.search(pattern, shape.text)
            
            if match:
                account_name = match.group(1).strip()
                message_voom = match.group(2).strip()
                if re.search(r"メッセージ",message_voom):
                    message_or_voom = 1
                elif re.search(r"VOOM",message_voom):
                    message_or_voom = 2
        elif check_position(shape,normal_permissible,cover_position_top[1],cover_position_left[1]):
            # 正規表現パターンを定義
            pattern = r"(\d{1,2})月(\d{1,2})日.*?(\d{1,2}:\d{2})"
            
            # パターンにマッチするテキストを検索
            match = re.search(pattern, shape.text)
            
            if match:
                month = int(match.group(1))
                day = int(match.group(2))
                time = match.group(3)
        #1.LPとADの大まかな範囲を指定(top厳密,left指定なし)
        #2.幅でpictureをLPとADを分類、位置判定、数
        #3.振り数を分類,位置判定,数
        #3.矢印有無
        elif (shape.top.pt>reference_line_top[0] - strict_permissible):
            if (shape.shape_type == 13, shape.width ==  102.0455905511811):
                if check_position(shape,strict_permissible,reference_line_top[1],reference_line_left[1]):
                    ad_presence = True
                    if shape.top.pt +shape.height.pt > 145:
                        error_message += "adの高さが大きすぎる,"
                elif check_position(shape,strict_permissible,reference_line_top[1],reference_line_left[2]):
                    ad2_bottom = shape.top.pt +shape.height.pt
                else:
                    error_message += "基準線に従っていないAD,"
            elif (shape.shape_type == 13, shape.width ==  93.54181102362205):
                lp_count += 1
            elif (shape.auto_shape_type == 7):
                arrow_presence = shape.top.pt
            elif(shape.auto_shape_type == 1 and shape.left.pt > reference_line_left[3]-1):
                lp_number_count += 1
            elif(shape.auto_shape_type == 1 ):
                ad_number_count = +1
            else:
                error_message +="基準線にあっていないオブジェクトがあります。"
    if not(account_name and message_or_voom):
        error_message = "オブジェクトが基準値より20pt離れている"
    print(account_name,message_or_voom,year,month,error_message)


    # return pd.DataFrame([{
    #     'category_number': 3,
    #     'account_name': account_name,
    #     'message_or_voom': message_or_voom,
    #     'month': month,
    #     'day': day,
    #     'time': time,
    #     'ad_presence': ad_presence,
    #     'ad_account_name': ad_account_name,
    #     'lp_count': lp_count,
    #     'lp_number_count': lp_number_count,
    #     'arrow_presence': arrow_presence,
    #     'error_message': error_message
    # }])

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
            # df = 
            extract_cover_data(slide)
            print(0)
        elif slide_type == 'month':
            # df = 
            extract_month_data(slide)
            print(1)
        elif slide_type == 'content':
            # df = extract_content_data(slide)
            None
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

