import tkinter as tk
from tkinter import filedialog, messagebox
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from check import all_checks
import os
import pandas as pd

def save_df_to_excel(df, file_path):
    # ファイルが存在するかを確認
    if os.path.exists("check_message.xlsx"):
        # 既存のワークブックを読み込む
        book = openpyxl.load_workbook("check_message.xlsx")
    else:
        # 新しいワークブックを作成
        book = openpyxl.Workbook()
        # 新しいワークブックにデフォルトで作成されるシートを削除
        book.remove(book.active)
    
    # ファイル名から拡張子を除いた部分をシート名として抽出
    sheet_name = os.path.splitext(os.path.basename(file_path))[0]
    
    # ファイル名をシート名として新しいシートを作成
    sheet = book.create_sheet(title=sheet_name)
    
    # ファイル名を一行目に書き込む
    sheet.append([file_path])
    
    # データフレームの内容をシートに書き込む
    for r in dataframe_to_rows(df, index=True, header=True):
        sheet.append(r)
    
    # ワークブックを保存
    book.save("check_message.xlsx")

def open_files():
    # ファイルダイアログを開いてファイルを選択
    file_paths = filedialog.askopenfilenames(filetypes=[("PowerPoint files", "*.pptx")])
    
    # 選択された各ファイルに対して処理を実行
    for file_path in file_paths:
        df = all_checks(file_path)
        if not df.empty:
            messagebox.showinfo("結果", "ミスが検出されました。")
            save_df_to_excel(df, file_path)
        else:
            messagebox.showinfo("結果", "ミスは検出されませんでした。")

def show_message_screen():
    # メッセージ画面の作成
    for widget in window.winfo_children():
        widget.destroy()
    
    check_button = tk.Button(window, text="Check", command=open_files)
    check_button.pack()
    
    extract_button = tk.Button(window, text="抽出 (今後実装予定)")
    extract_button.pack()

def create_gui():
    global window
    # メインウィンドウの作成
    window = tk.Tk()
    window.title("PowerPoint Text Extractor")
    
    # 最初の画面のボタンを作成
    message_button = tk.Button(window, text="Message", command=show_message_screen)
    message_button.pack()
    
    voom_button = tk.Button(window, text="VOOM (今後実装予定)")
    voom_button.pack()
    
    window.mainloop()
