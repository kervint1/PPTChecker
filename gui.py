import tkinter as tk
from tkinter import filedialog, messagebox
from index import extract_text_from_pptx_by_slide,save_texts_to_excel
import os

def create_gui():
    window = tk.Tk()
    window.title("PowerPoint Text Extractor")

    def open_files():
        filepaths = filedialog.askopenfilenames(title="Select PowerPoint files", filetypes=[("PowerPoint files", "*.pptx")])
        all_texts = [(os.path.basename(filepath), extract_text_from_pptx_by_slide(filepath)) for filepath in filepaths]
        if all_texts:
            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
            if save_path:
                save_texts_to_excel(all_texts, save_path)
                messagebox.showinfo("Success", "Texts have been successfully saved to Excel.")

    open_button = tk.Button(window, text="Open PowerPoint Files and Save to Excel", command=open_files)
    open_button.pack()

    window.mainloop()
