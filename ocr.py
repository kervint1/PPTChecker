import pytesseract
from pptx import Presentation
from PIL import Image
import io
import os

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
os.environ['TESSDATA_PREFIX'] = r'C:\Program Files\Tesseract-OCR\tessdata'

def get_lp_account_name_message(file_path, left=55, top=10, right=200, bottom=100):
    prs = Presentation(file_path)
    slides_texts = []  # 各スライドのテキストリストを格納するリスト

    for slide in prs.slides:
        slide_texts = []  # 現在のスライドのテキストを格納するリスト
        for shape in slide.shapes:
            if shape.shape_type == 13 and shape.left.pt<37 and shape.left.pt>32 and shape.top.pt<147 and shape.top.pt>143:  # 画像タイプ
                image_stream = io.BytesIO(shape.image.blob)
                image = Image.open(image_stream)
                # 画像を指定された範囲でクロップ
                cropped_image = image.crop((left, top, right, bottom))
                text = pytesseract.image_to_string(cropped_image, lang='eng+jpn')
                slide_texts.append(text)
        slides_texts.append(slide_texts)

    return slides_texts

def get_lp_date_message(file_path, left=0, top=0, right=2000, bottom=2000):
    prs = Presentation(file_path)
    slides_texts = []  # 各スライドのテキストリストを格納するリスト

    for slide in prs.slides:
        slide_texts = []  # 現在のスライドのテキストを格納するリスト
        for shape in slide.shapes:
            if shape.shape_type == 13 and 32 < shape.left.pt < 37 and 143 < shape.top.pt < 147:  # 画像タイプ
                image_stream = io.BytesIO(shape.image.blob)
                image = Image.open(image_stream)
                # 画像を指定された範囲でクロップ
                cropped_image = image.crop((left, top, right, bottom))
                text = pytesseract.image_to_string(cropped_image, lang='jpn')
                slide_texts.append(text)
        slides_texts.append(slide_texts)

    return slides_texts

# ファイルパスを指定して関数を呼び出し、結果を表示
file_path = r"【事例資料】LOUIS VUITTON_LINE 公式アカウント_メッセージ配信_2024年1月以降.pptx"
print(get_lp_date_message(file_path))
