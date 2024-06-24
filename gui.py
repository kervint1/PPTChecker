import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from pptx import Presentation
import pandas as pd
from check import all_checks
from index import summarize_latest_slides
import os
import analysis

def open_files():
    file_paths = filedialog.askopenfilenames(filetypes=[("PowerPoint files", "*.pptx")])
    return file_paths

def process_files(file_paths, ocr, check, months, output):
    for file_path in file_paths:
        df = summarize_latest_slides(file_path, ocr, months)
        messages = []

        if check:
            messages = all_checks(df)
            if not messages:
                messages.append("ミスなし")

        if output:
            save_to_excel_csv(df, file_path)

        show_messages(file_path, messages)

def show_messages(file_path, messages):
    result = f"ファイル: {os.path.basename(file_path)}\n" + "\n".join(messages)
    messagebox.showinfo("チェック結果", result)

def save_to_excel_csv(df, file_path):
    base_name = os.path.splitext(os.path.basename(file_path))[0]
    save_dir = os.path.dirname(file_path)
    
    csv_dir = os.path.join(save_dir, "LINEcsv")
    excel_dir = os.path.join(save_dir, "LINEExcel")

    os.makedirs(csv_dir, exist_ok=True)
    os.makedirs(excel_dir, exist_ok=True)

    csv_path = os.path.join(csv_dir, f"{base_name}.csv")
    excel_path = os.path.join(excel_dir, f"{base_name}.xlsx")
    
    df.to_csv(csv_path, index=False)

    workbook = openpyxl.Workbook()
    sheet = workbook.active

    for r_idx, row in enumerate(dataframe_to_rows(df, index=True, header=True), 1):
        for c_idx, value in enumerate(row, 1):
            sheet.cell(row=r_idx, column=c_idx, value=value)
    
    workbook.save(excel_path)

def confirm_and_process():
    ocr = ocr_var.get()
    check = check_var.get()
    months = months_var.get()
    output = output_var.get()
    
    try:
        months = int(months) if months else None
    except ValueError:
        messagebox.showerror("入力エラー", "月数は整数で入力してください")
        return

    file_paths = open_files()
    if file_paths:
        process_files(file_paths, ocr, check, months, output)

def create_gui():
    global ocr_var, check_var, months_var, output_var

    # GUI作成
    root = tk.Tk()
    root.title("PPT Checker")

    frame = ttk.Frame(root, padding="10")
    frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

    ocr_var = tk.BooleanVar()
    check_var = tk.BooleanVar()
    months_var = tk.StringVar()
    output_var = tk.BooleanVar()

    ttk.Checkbutton(frame, text="OCR有無", variable=ocr_var).grid(row=0, column=0, sticky=tk.W)
    ttk.Checkbutton(frame, text="チェック有無", variable=check_var).grid(row=1, column=0, sticky=tk.W)
    ttk.Label(frame, text="何か月分:").grid(row=2, column=0, sticky=tk.W)
    ttk.Entry(frame, textvariable=months_var).grid(row=2, column=1, sticky=(tk.W, tk.E))
    ttk.Checkbutton(frame, text="データ出力有無", variable=output_var).grid(row=3, column=0, sticky=tk.W)

    ttk.Button(frame, text="確認", command=confirm_and_process).grid(row=4, column=0, columnspan=2, pady=10)
    ttk.Button(frame, text="分析", command=analysis.create_analysis_gui).grid(row=5, column=0, columnspan=2, pady=10)

    root.mainloop()
