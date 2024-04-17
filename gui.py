import tkinter as tk
from tkinter import filedialog
from index import extract_text_from_pptx_by_slide

def create_gui():
    window = tk.Tk()
    window.title("PowerPoint Text Extractor")

    def open_file():
        filepath = filedialog.askopenfilename()
        if not filepath:
            return
        texts = extract_text_from_pptx_by_slide(filepath)
        text_area.delete("1.0", tk.END)
        text_area.insert(tk.END, "\n".join(str(texts[0][0])))

    open_button = tk.Button(window, text="Open PowerPoint File", command=open_file)
    open_button.pack()

    text_area = tk.Text(window, height=10, width=50)
    text_area.pack()

    window.mainloop()
