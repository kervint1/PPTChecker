import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import openpyxl
import os
from statistics import mode

def open_csv_files():
    file_paths = filedialog.askopenfilenames(filetypes=[("CSV files", "*.csv")])
    return file_paths

def analyze_files(file_paths):
    ad_data = []
    lp_data = []
    
    for file_path in file_paths:
        df = pd.read_csv(file_path, parse_dates=['date'])
        file_name = os.path.basename(file_path).split('.')[0]
        
        monthly_ad_count = df[df['category_number'] == 3].groupby(df['date'].dt.month).size()
        monthly_lp_avg = df[df['category_number'] == 3].groupby(df['date'].dt.month)['ad_number_count'].mean()
        
        ad_row = [file_name] + [monthly_ad_count.get(i, 0) for i in range(1, 13)]
        lp_row = [file_name] + [round(monthly_lp_avg.get(i, 0), 2) for i in range(1, 13)]
        
        ad_total = sum(ad_row[1:])
        lp_total = sum(lp_row[1:])
        
        ad_row += [
            ad_total,
            round(ad_total / 12, 2),
            mode(ad_row[1:13]) if ad_row[1:13] else 0,
            max(ad_row[1:13]) if ad_row[1:13] else 0
        ]
        
        lp_row += [
            lp_total,
            round(lp_total / 12, 2),
            mode(lp_row[1:13]) if lp_row[1:13] else 0,
            max(lp_row[1:13]) if lp_row[1:13] else 0
        ]
        
        ad_data.append(ad_row)
        lp_data.append(lp_row)
    
    save_analysis_to_excel(ad_data, lp_data)

def save_analysis_to_excel(ad_data, lp_data):
    workbook = openpyxl.Workbook()

    # 月ごとAD数シート
    ad_sheet = workbook.active
    ad_sheet.title = '月ごとAD数'
    ad_header = ["ファイル名"] + [f"{i}月" for i in range(1, 13)] + ["合計", "平均", "最頻値", "最大値"]
    ad_sheet.append(ad_header)
    for row in ad_data:
        ad_sheet.append(row)

    # 月ごとLP数シート
    lp_sheet = workbook.create_sheet(title='月ごと平均LP数')
    lp_header = ["ファイル名"] + [f"{i}月" for i in range(1, 13)] + ["合計", "平均", "最頻値", "最大値"]
    lp_sheet.append(lp_header)
    for row in lp_data:
        lp_sheet.append(row)

    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        workbook.save(file_path)
        messagebox.showinfo("完了", f"分析結果が {file_path} に保存されました")

def create_analysis_gui():
    root = tk.Tk()
    root.withdraw()  # メインウィンドウを表示しない
    file_paths = open_csv_files()
    if file_paths:
        analyze_files(file_paths)
    root.destroy()

if __name__ == "__main__":
    create_analysis_gui()
