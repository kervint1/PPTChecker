import os
from index import extract_text_from_pptx_by_slide

# 月の表紙スライドチェック
def monthCheck(filepaths):
    all_texts = [(os.path.basename(filepath), extract_text_from_pptx_by_slide(filepath)) for filepath in filepaths]
    