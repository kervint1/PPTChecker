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
