# -*- coding: utf-8 -*-
# GUI: Tkinter
# 要件: pypdf, pdf2image, pillow (poppler 依存)
# ---- 2025-05-13 update ----
import os
import re
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from pypdf import PdfReader, PdfWriter
from pdf2image import convert_from_path
from PIL import Image, ImageTk

PAGE_RE = re.compile(r'^\d+$')
RANGE_RE = re.compile(r'^(\d+)-(\d+)$')


def parse_pages(spec: str, max_page: int):
    """ページ指定文字列を解析し、0-index 昇順リストを返す。"""
    pages = set()
    spec = spec.strip()
    if not spec:
        return []
    for token in spec.split(','):
        token = token.strip()
        if not token:
            continue
        if PAGE_RE.match(token):
            num = int(token)
            if not (1 <= num <= max_page):
                raise ValueError(f'ページ {num} は 1〜{max_page} の範囲外')
            pages.add(num - 1)
        else:
            m = RANGE_RE.match(token)
            if not m:
                raise ValueError(f'書式エラー: {token}')
            a, b = map(int, m.groups())
            if a > b:
                raise ValueError(f'範囲 {token}: 開始 > 終了')
            if a < 1 or b > max_page:
                raise ValueError(f'範囲 {token} は 1〜{max_page} の範囲外')
            pages.update(range(a - 1, b))
    return sorted(pages)


class PDFExtractorApp:
    THUMB_DPI = 50        # 低解像度サムネイル
    THUMB_MAX_W = 140     # サムネイル最大幅 (px)
    COLOR_SEL   = '#ff6600'   # 選択枠 (オレンジ例)
    COLOR_NORM  = 'gray80'    # 非選択枠

    def __init__(self, root):
        self.root = root
        self.root.title('PDF ページ抽出 + サムネイル')

        # ---- 状態変数 ----
        self.pdf_path = tk.StringVar()
        self.page_spec = tk.StringVar()
        self.save_path = tk.StringVar()
        self.status = tk.StringVar(value='待機中')
        self.selected_pages = set()        # 0-index
        self.photo_images = []             # GC 防止用
        self.reader = None                 # PdfReader

        # ---- 左パネル ----
        left = ttk.Frame(root)
        left.grid(row=0, column=0, sticky='nsew')
        ttk.Button(left, text='PDF を選択', command=self.select_pdf)\
            .grid(row=0, column=0, padx=4, pady=4, sticky='ew')
        ttk.Entry(left, textvariable=self.pdf_path, width=45, state='readonly')\
            .grid(row=0, column=1, padx=4, pady=4, sticky='w')
        ttk.Label(left, text='ページ指定:')\
            .grid(row=1, column=0, padx=4, pady=4, sticky='e')
        ttk.Entry(left, textvariable=self.page_spec, width=30)\
            .grid(row=1, column=1, padx=4, pady=4, sticky='w')
        ttk.Button(left, text='保存先を選択', command=self.select_save)\
            .grid(row=2, column=0, padx=4, pady=4, sticky='ew')
        ttk.Entry(left, textvariable=self.save_path, width=45, state='readonly')\
            .grid(row=2, column=1, padx=4, pady=4, sticky='w')
        self.progress = ttk.Progressbar(left, length=380, mode='determinate')
        self.progress.grid(row=3, column=0, columnspan=2, padx=4, pady=6)
        ttk.Label(left, textvariable=self.status)\
            .grid(row=4, column=0, columnspan=2, pady=(0, 6))
        ttk.Button(left, text='実行', command=self.run)\
            .grid(row=5, column=0, columnspan=2, padx=4, pady=(0, 8), sticky='ew')

        # ---- 右パネル（サムネイル） ----
        thumb_frame = ttk.Frame(root, relief='sunken', borderwidth=1)
        thumb_frame.grid(row=0, column=1, sticky='nsew')
        self.canvas = tk.Canvas(thumb_frame, width=PDFExtractorApp.THUMB_MAX_W + 20)
        scrollbar = ttk.Scrollbar(thumb_frame, orient='vertical',
                                  command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=scrollbar.set)
        self.inner = ttk.Frame(self.canvas)
        self.canvas.create_window((0, 0), window=self.inner, anchor='nw')
        self.inner.bind('<Configure>',
                        lambda e: self.canvas.configure(scrollregion=self.canvas.bbox('all')))
        self.canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        # ---- レイアウト設定 ----
        root.grid_columnconfigure(0, weight=1)
        root.grid_columnconfigure(1, weight=0)
        root.grid_rowconfigure(0, weight=1)

    # ---------------------- UI コールバック ----------------------

    def select_pdf(self):
        path = filedialog.askopenfilename(title='PDF を選択', filetypes=[('PDF files', '*.pdf')])
        if not path:
            return
        self.pdf_path.set(path)
        self.status.set('PDF 読み込み中...')
        self.clear_thumbnails()
        self.root.update_idletasks()

        # 読み込みとサムネイル生成を別スレッドで
        threading.Thread(target=self.load_pdf_and_thumbnails, args=(path,), daemon=True).start()

    def select_save(self):
        path = filedialog.asksaveasfilename(title='保存先を指定',
                                            defaultextension='.pdf',
                                            filetypes=[('PDF files', '*.pdf')])
        if path:
            self.save_path.set(path)

    def run(self):
        if not self.reader:
            messagebox.showerror('エラー', '先に PDF を読み込んでください')
            return
        try:
            pages = parse_pages(self.page_spec.get(), len(self.reader.pages))
            if not pages:
                raise ValueError('抽出ページが指定されていません')
            save_file = self.save_path.get()
            if not save_file:
                raise ValueError('保存先が指定されていません')

            if os.path.exists(save_file):
                if not messagebox.askyesno('上書き確認', '既にファイルが存在します。上書きしますか？'):
                    return

            self.status.set('ページ抽出中...')
            self.progress['value'] = 0
            self.progress['maximum'] = len(pages)
            self.root.update_idletasks()

            writer = PdfWriter()
            for i, p in enumerate(pages, 1):
                writer.add_page(self.reader.pages[p])
                self.progress['value'] = i
                self.root.update_idletasks()

            if self.reader.metadata:
                writer.add_metadata(self.reader.metadata)
            with open(save_file, 'wb') as f:
                writer.write(f)

            self.status.set(f'抽出完了 ({len(pages)}ページ)')
            messagebox.showinfo('完了', 'PDF を保存しました')
        except Exception as e:
            self.status.set('失敗')
            messagebox.showerror('エラー', str(e))
        finally:
            self.progress['value'] = 0

    # ---------------------- PDF & Thumbnail ----------------------

    def load_pdf_and_thumbnails(self, path):
        try:
            reader = PdfReader(path)
            if reader.is_encrypted:
                raise RuntimeError('暗号化された PDF は扱えません')
            self.reader = reader
            total_pages = len(reader.pages)

            # ----- サムネイル生成 -----
            self.progress['maximum'] = total_pages
            self.progress['value'] = 0
            self.status.set('サムネイル生成中...')
            self.root.update_idletasks()

            images = convert_from_path(path, dpi=PDFExtractorApp.THUMB_DPI,
                                       fmt='png', thread_count=2)
            thumbs = []
            for idx, img in enumerate(images):
                # 縮小 (幅 140px 未満ならそのまま)
                w, h = img.size
                if w > PDFExtractorApp.THUMB_MAX_W:
                    ratio = PDFExtractorApp.THUMB_MAX_W / w
                    img = img.resize((int(w * ratio), int(h * ratio)), Image.BILINEAR)
                thumbs.append(img)
                self.progress['value'] = idx + 1
                self.root.update_idletasks()

            self.photo_images = [ImageTk.PhotoImage(t) for t in thumbs]

            # Canvas に貼り付け（メインスレッドで）
            self.root.after(0, self.populate_thumbnail_canvas)

            self.status.set('PDF 読み込み完了')
        except Exception as e:
            self.status.set('失敗')
            messagebox.showerror('エラー', str(e))
        finally:
            self.progress['value'] = 0

    def populate_thumbnail_canvas(self):
        for widget in self.inner.winfo_children():
            widget.destroy()
        self.selected_pages.clear()

        for i, photo in enumerate(self.photo_images, 1):
            frm = tk.Frame(self.inner,  # ← tk.Frame に
                        bd=0,
                        highlightthickness=2,           # 枠線太さ
                        highlightbackground='gray80')   # 非選択時の色
            lbl = ttk.Label(frm, image=photo)
            lbl.pack()
            cap = ttk.Label(frm, text=f'{i}', width=4, anchor='e')
            cap.pack()
            frm.grid(row=i - 1, column=0, sticky='n', padx=2, pady=2)
            # クリックバインド
            for w in (frm, lbl, cap):
                w.bind('<Button-1>', lambda e, page=i: self.toggle_page(page))

    def clear_thumbnails(self):
        for w in self.inner.winfo_children():
            w.destroy()
        self.photo_images.clear()
        self.selected_pages.clear()

    # ---------------------- ページ選択連携 ----------------------

    def toggle_page(self, page_number):
        """1-index page_number を選択/解除し見た目と Entry を同期"""
        idx = page_number - 1
        frame = self.inner.grid_slaves(row=idx, column=0)[0]
        if idx in self.selected_pages:
            self.selected_pages.remove(idx)
            frame.configure(highlightbackground=PDFExtractorApp.COLOR_NORM)
        else:
            self.selected_pages.add(idx)
            frame.configure(highlightbackground=PDFExtractorApp.COLOR_SEL)
        self.update_page_spec_entry()

    def update_page_spec_entry(self):
        if not self.selected_pages:
            self.page_spec.set('')
            return
        pages = sorted(self.selected_pages)             # 0-index
        ranges = []
        start = prev = pages[0]
        for p in pages[1:]:
            if p == prev + 1:
                prev = p
            else:
                ranges.append((start, prev))
                start = prev = p
        ranges.append((start, prev))
        spec_parts = [f'{a+1}' if a == b else f'{a+1}-{b+1}' for a, b in ranges]
        self.page_spec.set(','.join(spec_parts))


# ---- スタイル（選択ハイライト枠色） ----
def setup_styles():
    style = ttk.Style()
    style.configure('Selected.TFrame', borderwidth=4, relief='solid')


def main():
    root = tk.Tk()
    setup_styles()
    app = PDFExtractorApp(root)
    root.mainloop()


if __name__ == '__main__':
    main()
