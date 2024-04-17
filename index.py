from pptx import Presentation

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

# ファイルパスを指定して関数を呼び出し、結果を表示します。
print(extract_text_from_pptx_by_slide(r"C:\Users\iniad\Documents\adpro\テスト4月4日.pptx"))
