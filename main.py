import tkinter as tk
from tkinter import filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
import os

from PyPDF2 import PdfMerger, PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import io


def add_page_numbers(input_pdf, output_pdf):
    """PDFにページ番号を追加"""
    reader = PdfReader(input_pdf)
    writer = PdfWriter()

    for i, page in enumerate(reader.pages, start=1):
        packet = io.BytesIO()
        can = canvas.Canvas(packet, pagesize=letter)
        can.setFont("Helvetica", 10)
        can.drawString(500, 20, f"{i}")  # ページ番号（右下）
        can.save()

        packet.seek(0)
        overlay_pdf = PdfReader(packet)
        page.merge_page(overlay_pdf.pages[0])
        writer.add_page(page)

    with open(output_pdf, "wb") as f:
        writer.write(f)


class PDFMergerApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF結合簡易ソフトウェア")
        self.geometry("400x300")

        self.file_listbox = tk.Listbox(self, selectmode=tk.SINGLE)
        self.file_listbox.pack(fill=tk.BOTH, expand=True)

        # ドラッグアンドドロップ受付
        self.file_listbox.drop_target_register(DND_FILES)
        self.file_listbox.dnd_bind("<<Drop>>", self.drop_files)

        # チェックボックス（ページ番号振るかどうか）
        self.add_numbers = tk.BooleanVar(value=True)  # デフォルトON
        chk = tk.Checkbutton(self, text="ページ番号を付与する", variable=self.add_numbers)
        chk.pack(anchor="w", padx=10, pady=5)

        # ボタン類
        btn_frame = tk.Frame(self)
        btn_frame.pack(fill=tk.X)

        tk.Button(btn_frame, text="↑ 上へ", command=self.move_up).pack(side=tk.LEFT)
        tk.Button(btn_frame, text="↓ 下へ", command=self.move_down).pack(side=tk.LEFT)
        tk.Button(btn_frame, text="削除", command=self.delete_selected).pack(side=tk.LEFT)
        tk.Button(btn_frame, text="結合", command=self.merge_pdfs).pack(side=tk.RIGHT)

    def drop_files(self, event):
        files = self.split_files(event.data)
        for f in files:
            if f.lower().endswith(".pdf"):
                self.file_listbox.insert(tk.END, f)

    def split_files(self, data):
        # Windows / macOS のファイルパス分割対応
        return self.tk.splitlist(data)

    def move_up(self):
        idx = self.file_listbox.curselection()
        if not idx:
            return
        i = idx[0]
        if i == 0:
            return
        text = self.file_listbox.get(i)
        self.file_listbox.delete(i)
        self.file_listbox.insert(i - 1, text)
        self.file_listbox.select_set(i - 1)

    def move_down(self):
        idx = self.file_listbox.curselection()
        if not idx:
            return
        i = idx[0]
        if i == self.file_listbox.size() - 1:
            return
        text = self.file_listbox.get(i)
        self.file_listbox.delete(i)
        self.file_listbox.insert(i + 1, text)
        self.file_listbox.select_set(i + 1)

    def delete_selected(self):
        idx = self.file_listbox.curselection()
        if not idx:
            return
        self.file_listbox.delete(idx[0])

    def merge_pdfs(self):
        files = self.file_listbox.get(0, tk.END)
        if not files:
            messagebox.showerror("Error", "エラーが発生しました")
            return

        save_path = filedialog.asksaveasfilename(defaultextension=".pdf",
                                                 filetypes=[("PDF files", "*.pdf")])
        if not save_path:
            return

        # 一旦結合
        merger = PdfMerger()
        for pdf in files:
            merger.append(pdf)

        tmp_path = save_path.replace(".pdf", "_tmp.pdf")
        with open(tmp_path, "wb") as f:
            merger.write(f)
        merger.close()

        # ページ番号付与はチェック次第
        if self.add_numbers.get():
            add_page_numbers(tmp_path, save_path)
            os.remove(tmp_path)
        else:
            os.rename(tmp_path, save_path)

        messagebox.showinfo("Success", f"結合したPDが保存されました:\n{save_path}")


if __name__ == "__main__":
    app = PDFMergerApp()
    app.mainloop()
