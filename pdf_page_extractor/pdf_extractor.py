# --- 使用ライブラリ ---
# GUI構築用: Tkinter
# PDF読み書き: pypdf
# PDF → 画像変換 (サムネイル生成用): pdf2image （※Poppler 必須）
# 画像処理 & GUI表示用: Pillow
import os
import re
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from pypdf import PdfReader, PdfWriter
from pdf2image import convert_from_path # サムネイル機能に必要
from PIL import Image, ImageTk # サムネイル機能に必要

PAGE_RE = re.compile(r'^\d+$')
RANGE_RE = re.compile(r'^(\d+)-(\d+)$')


def parse_pages(spec: str, max_page: int):
    """
    ユーザーの入力（例: "1,3-5"）を解析し、
    0始まりのページ番号リスト（昇順）を返す。
    入力値のバリデーションも同時に行う。
    """
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
        self.split_pages_var = tk.StringVar()  # "Pページ毎に分割"機能で使用される分割サイズ(P)の入力用

        # ---- 左パネル ----
        left = ttk.Frame(root)
        left.grid(row=0, column=0, sticky='nsew', padx=5, pady=5)

        # PDF選択
        ttk.Button(left, text='PDF を選択', command=self.select_pdf)\
            .grid(row=0, column=0, padx=4, pady=4, sticky='ew')
        ttk.Entry(left, textvariable=self.pdf_path, width=45, state='readonly')\
            .grid(row=0, column=1, padx=4, pady=4, sticky='w')

        # ページ指定 (抽出用)
        ttk.Label(left, text='ページ指定:')\
            .grid(row=1, column=0, padx=4, pady=4, sticky='e')
        ttk.Entry(left, textvariable=self.page_spec, width=30)\
            .grid(row=1, column=1, padx=4, pady=4, sticky='w')

        # 保存先選択 (抽出用、および分割時のベースとして使用)
        ttk.Button(left, text='保存先を選択', command=self.select_save)\
            .grid(row=2, column=0, padx=4, pady=4, sticky='ew')
        ttk.Entry(left, textvariable=self.save_path, width=45, state='readonly')\
            .grid(row=2, column=1, padx=4, pady=4, sticky='w')

        # プログレスバーとステータス (共用)
        self.progress = ttk.Progressbar(left, length=380, mode='determinate')
        self.progress.grid(row=3, column=0, columnspan=2, padx=4, pady=6, sticky='ew')
        ttk.Label(left, textvariable=self.status)\
            .grid(row=4, column=0, columnspan=2, pady=(0, 6))

        # ページ指定抽出実行ボタン
        ttk.Button(left, text='ページ指定抽出実行', command=self.run)\
            .grid(row=5, column=0, columnspan=2, padx=4, pady=(4, 8), sticky='ew')
        
        # セパレーター (視覚的な区切り)
        ttk.Separator(left, orient='horizontal').grid(row=6, column=0, columnspan=2, sticky='ew', pady=10)

        # Pページ毎に分割 機能
        ttk.Label(left, text='Pページ毎に分割:')\
            .grid(row=7, column=0, padx=4, pady=4, sticky='e')
        ttk.Entry(left, textvariable=self.split_pages_var, width=10)\
            .grid(row=7, column=1, padx=4, pady=4, sticky='w')
        ttk.Button(left, text='分割実行', command=self.run_split)\
            .grid(row=8, column=0, columnspan=2, padx=4, pady=(4, 8), sticky='ew')

        # ---- 右パネル（サムネイル） ----
        thumb_frame = ttk.Frame(root, relief='sunken', borderwidth=1)
        thumb_frame.grid(row=0, column=1, sticky='nsew', padx=(0,5), pady=5)
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
        root.grid_columnconfigure(1, weight=0) # サムネイルパネルの幅は固定
        root.grid_rowconfigure(0, weight=1)

    # ---------------------- UI コールバック ----------------------

    def select_pdf(self):
        path = filedialog.askopenfilename(title='PDF を選択', filetypes=[('PDF files', '*.pdf')])
        if not path:
            return # ユーザーがキャンセルした場合は処理を中止
        self.pdf_path.set(path)
        self.status.set('PDF 読み込み中...')
        self.clear_thumbnails()
        self.root.update_idletasks()

        threading.Thread(target=self.load_pdf_and_thumbnails, args=(path,), daemon=True).start()

    def select_save(self):
        path = filedialog.asksaveasfilename(title='保存先を指定',
                                            defaultextension='.pdf',
                                            filetypes=[('PDF files', '*.pdf')])
        if path:
            self.save_path.set(path)

    def run(self):
        """指定ページ抽出を実行"""
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
            for i, p_idx in enumerate(pages, 1): # p_idx は 0-indexed
                writer.add_page(self.reader.pages[p_idx])
                self.progress['value'] = i
                self.root.update_idletasks()

            if self.reader.metadata:
                writer.add_metadata(self.reader.metadata)
            with open(save_file, 'wb') as f:
                writer.write(f)

            self.status.set(f'抽出完了 ({len(pages)}ページ)')
            messagebox.showinfo('完了', 'PDF を保存しました')
        except Exception as e:
            self.status.set('抽出失敗')
            messagebox.showerror('エラー', str(e))
        finally:
            self.progress['value'] = 0
            
    def run_split(self):
        """Pページ毎に分割を実行"""
        if not self.reader:
            messagebox.showerror('エラー', '先に PDF を読み込んでください')
            return

        try:
            p_str = self.split_pages_var.get().strip()
            if not p_str.isdigit() or int(p_str) <= 0:
                raise ValueError('分割ページ数 P は正の整数で指定してください')
            p_val = int(p_str)

            save_file_spec = self.save_path.get()
            pdf_full_path = self.pdf_path.get()
            pdf_dir_fallback = os.path.dirname(pdf_full_path)
            pdf_basename_fallback = os.path.splitext(os.path.basename(pdf_full_path))[0]

            output_dir: str
            base_name_for_split: str

            if not save_file_spec:
                msg = (f"保存先が指定されていません。\n"
                       f"元のPDFと同じディレクトリ ({pdf_dir_fallback}) に \n"
                       f"'{pdf_basename_fallback}_pages_X-Y.pdf' の形式で保存しますか？\n"
                       f"（キャンセルすると処理を中断します）")
                if not messagebox.askokcancel("確認", msg):
                    self.status.set('保存先を指定してください')
                    self.progress['value'] = 0
                    return
                output_dir = pdf_dir_fallback
                base_name_for_split = pdf_basename_fallback
            else:
                output_dir = os.path.dirname(save_file_spec)
                base_name_for_split = os.path.splitext(os.path.basename(save_file_spec))[0]
                if not output_dir: # ファイル名のみが指定された場合 (例: "output.pdf")
                    output_dir = pdf_dir_fallback 
            
            if not os.path.exists(output_dir) and output_dir : # output_dirが空文字でないことを確認
                try:
                    os.makedirs(output_dir, exist_ok=True)
                except OSError as e:
                    raise ValueError(f"出力ディレクトリの作成に失敗しました: {output_dir}\n{e}")
            elif not output_dir : # output_dirが空文字だった場合（基本的には発生しないはずだが念のため）
                output_dir = "." # カレントディレクトリ


            total_pages = len(self.reader.pages)
            if total_pages == 0:
                raise ValueError('PDFにページがありません')

            num_output_files = (total_pages + p_val - 1) // p_val

            self.status.set(f'{p_val}ページ毎に分割中...')
            self.progress['value'] = 0
            self.progress['maximum'] = num_output_files
            self.root.update_idletasks()

            generated_files_count = 0
            for i in range(num_output_files):
                writer = PdfWriter()
                start_page_idx_0based = i * p_val
                end_page_idx_0based = min((i + 1) * p_val, total_pages)

                for page_num_0based in range(start_page_idx_0based, end_page_idx_0based):
                    writer.add_page(self.reader.pages[page_num_0based])

                if self.reader.metadata:
                    writer.add_metadata(self.reader.metadata)
                
                # 出力ファイル名 (ページ番号は1-indexedで表示)
                start_page_1based = start_page_idx_0based + 1
                end_page_1based = end_page_idx_0based
                
                if p_val == 1:
                    output_filename = os.path.join(output_dir, f"{base_name_for_split}_page_{start_page_1based}.pdf")
                else:
                    output_filename = os.path.join(output_dir, f"{base_name_for_split}_pages_{start_page_1based}-{end_page_1based}.pdf")

                if os.path.exists(output_filename):
                    if not messagebox.askyesno('上書き確認', f'ファイル "{os.path.basename(output_filename)}" は既に存在します。上書きしますか？'):
                        self.status.set(f'ファイル "{os.path.basename(output_filename)}" の上書きをスキップしました。')
                        self.progress['value'] += 1 # プログレスは進める
                        self.root.update_idletasks()
                        continue 

                with open(output_filename, 'wb') as f:
                    writer.write(f)
                generated_files_count += 1
                self.progress['value'] +=1
                self.root.update_idletasks()

            self.status.set(f'分割完了 ({generated_files_count}個のファイルを出力)')
            if generated_files_count > 0:
                messagebox.showinfo('完了', f'{generated_files_count}個のPDFファイルを下記に保存しました:\n{os.path.abspath(output_dir)}')
            else:
                messagebox.showinfo('情報', '出力するファイルはありませんでした（すべてスキップされたか、対象ページが0でした）。')

        except Exception as e:
            self.status.set('分割失敗')
            messagebox.showerror('エラー', str(e))
        finally:
            self.progress['value'] = 0

    # ---------------------- PDF & Thumbnail ----------------------

    def load_pdf_and_thumbnails(self, path):
        try:
            reader = PdfReader(path)
            if reader.is_encrypted:
                # self.reader = None # readerをクリアする
                raise RuntimeError('暗号化された PDF は扱えません')
            self.reader = reader
            total_pages = len(reader.pages)
            if total_pages == 0:
                self.status.set('PDF 読み込み完了 (0ページ)')
                self.progress['value'] = 0
                # サムネイル処理は不要
                self.root.after(0, self.populate_thumbnail_canvas) # キャンバスをクリアするために呼ぶ
                return

            self.progress['maximum'] = total_pages
            self.progress['value'] = 0
            self.status.set('サムネイル生成中...')
            self.root.update_idletasks()

            # pdf2image が Poppler に依存
            images = convert_from_path(path, dpi=PDFExtractorApp.THUMB_DPI,
                                       fmt='png', thread_count=2, first_page=1, last_page=min(total_pages, 100)) # 全ページのサムネイル生成は負荷が高いため、最大100ページまでに制限
            
            thumbs = []
            for idx, img in enumerate(images):
                w, h = img.size
                if w > PDFExtractorApp.THUMB_MAX_W:
                    ratio = PDFExtractorApp.THUMB_MAX_W / w
                    img = img.resize((int(w * ratio), int(h * ratio)), Image.Resampling.BILINEAR) # Pillow 9.0.0+
                thumbs.append(img)
                # サムネイル生成の進捗表示は、実際の画像処理枚数に合わせる
                self.progress['value'] = idx + 1
                if idx + 1 >= self.progress['maximum']: # convert_from_path でページ制限した場合
                    self.progress['maximum'] = idx + 1 
                self.root.update_idletasks()
            
            self.photo_images = [ImageTk.PhotoImage(t) for t in thumbs]
            self.root.after(0, self.populate_thumbnail_canvas)
            self.status.set(f'PDF 読み込み完了 ({total_pages}ページ)')
        except RuntimeError as e: # 特に暗号化PDFなど
             self.status.set('PDF読み込み失敗')
             self.reader = None # readerをクリア
             messagebox.showerror('PDF読み込みエラー', str(e))
        except ImportError:
            self.status.set('サムネイル機能エラー (pdf2image/Pillow?)')
            messagebox.showwarning('警告', 'サムネイル表示に必要なライブラリ (pdf2image, Pillow) または依存関係 (Poppler) が見つからない可能性があります。')
            # サムネイルなしでも抽出/分割は可能なように reader は設定しておく
            if 'reader' not in locals() or not isinstance(reader, PdfReader): # readerが未設定の場合のみ
                try:
                    self.reader = PdfReader(path) # サムネイルなしでリーダーだけ初期化試行
                    if self.reader.is_encrypted:
                         self.status.set('暗号化PDFは扱えません')
                         self.reader = None
                         messagebox.showerror('エラー','暗号化されたPDFは扱えません。')
                         return
                    self.status.set(f'PDF読込完了(サムネイルなし) ({len(self.reader.pages)}ページ)')
                except Exception as e_read:
                    self.status.set('PDF読み込み失敗')
                    self.reader = None
                    messagebox.showerror('PDF読み込みエラー', f"PDFの読み込みに失敗しました: {e_read}")
        except Exception as e:
            self.status.set('失敗')
            self.reader = None # readerをクリア
            messagebox.showerror('エラー', str(e))
        finally:
            self.progress['value'] = 0


    def populate_thumbnail_canvas(self):
        for widget in self.inner.winfo_children():
            widget.destroy()
        # self.selected_pages.clear() # サムネイル再描画時に選択をクリアするかは仕様による。現状は維持。

        if not self.photo_images: # サムネイルがない場合は何もしないか、メッセージ表示
            if self.reader and len(self.reader.pages) > 0: # PDF自体は読み込めている場合
                 # ttk.Label(self.inner, text="サムネイル表示不可").pack(padx=10, pady=10)
                 # ページ番号だけでも表示する (サムネイルなし版)
                 for i in range(1, len(self.reader.pages) + 1):
                    frm = tk.Frame(self.inner, bd=0, highlightthickness=2,
                                  highlightbackground=(PDFExtractorApp.COLOR_SEL if (i-1) in self.selected_pages else PDFExtractorApp.COLOR_NORM))
                    # ダミーのラベルでページ番号を表示
                    lbl_text = f"ページ {i}"
                    lbl = ttk.Label(frm, text=lbl_text, width=int(PDFExtractorApp.THUMB_MAX_W / 7), anchor='center') # 幅を調整
                    lbl.pack(pady=5)
                    # cap = ttk.Label(frm, text=f'{i}', width=4, anchor='e') # ページ番号は上記に含めた
                    # cap.pack()
                    frm.grid(row=i - 1, column=0, sticky='ew', padx=2, pady=2)
                    for w in (frm, lbl):
                        w.bind('<Button-1>', lambda e, page=i: self.toggle_page(page))
            return


        for i, photo in enumerate(self.photo_images, 1):
            # page_idx_0based = i -1 # このiはphoto_imagesのインデックスなので、必ずしもPDFページ番号ではない場合がある（サムネイル生成を途中までにした場合など）
            # サムネイルが表示されている実際のページ番号を使う必要がある。
            # 現状の実装ではphoto_imagesはPDFの先頭からの連番画像なので i がそのまま1-indexedページ番号として扱える。
            page_idx_0based = i - 1

            frm = tk.Frame(self.inner,
                        bd=0,
                        highlightthickness=2,
                        highlightbackground=(PDFExtractorApp.COLOR_SEL if page_idx_0based in self.selected_pages else PDFExtractorApp.COLOR_NORM))
            lbl = ttk.Label(frm, image=photo)
            lbl.pack()
            cap = ttk.Label(frm, text=f'{i}', width=4, anchor='e')
            cap.pack()
            frm.grid(row=page_idx_0based, column=0, sticky='n', padx=2, pady=2)
            for w in (frm, lbl, cap):
                w.bind('<Button-1>', lambda e, page_num=i: self.toggle_page(page_num))
        
        self.update_page_spec_entry() # サムネイル再描画後にも選択状態をEntryに反映


    def clear_thumbnails(self):
        for w in self.inner.winfo_children():
            w.destroy()
        self.photo_images.clear()
        # self.selected_pages.clear() # PDFが変わるので選択もクリアする
        # self.page_spec.set('')     # ページ指定Entryもクリア


    # ---------------------- ページ選択連携 ----------------------

    def toggle_page(self, page_number_1based):
        """1-index page_number を選択/解除し見た目と Entry を同期"""
        idx_0based = page_number_1based - 1
        
        # サムネイルリスト内の該当フレームを探す
        # grid_slaves はインデックスアクセスが不安定な場合があるので、子ウィジェットを直接たどる方が確実な場合もある
        try:
            # grid_slaves(row=idx_0based, column=0) は、そのrowに複数のウィジェットがある場合リストを返す
            # innerの子ウィジェットを調べて、該当するフレームを見つける
            target_frame = None
            for child_frame in self.inner.winfo_children():
                # child_frame (tk.Frame) の grid 情報から row を取得
                grid_info = child_frame.grid_info()
                if grid_info and grid_info.get('row') == idx_0based:
                    target_frame = child_frame
                    break
            
            if not target_frame: return #該当フレームが見つからない

        except tk.TclError: # grid_slaves が空の時にエラーになることがある
             return # 対象ウィジェットがない

        if idx_0based in self.selected_pages:
            self.selected_pages.remove(idx_0based)
            target_frame.configure(highlightbackground=PDFExtractorApp.COLOR_NORM)
        else:
            self.selected_pages.add(idx_0based)
            target_frame.configure(highlightbackground=PDFExtractorApp.COLOR_SEL)
        self.update_page_spec_entry()

    def update_page_spec_entry(self):
        if not self.selected_pages:
            self.page_spec.set('')
            return
        # selected_pages は 0-index なので、1-index に変換して表示
        pages_1based = sorted([p + 1 for p in self.selected_pages])
        
        if not pages_1based:
            self.page_spec.set('')
            return

        ranges = []
        start = prev = pages_1based[0]
        for p_1based in pages_1based[1:]:
            if p_1based == prev + 1:
                prev = p_1based
            else:
                ranges.append((start, prev))
                start = prev = p_1based
        ranges.append((start, prev))
        
        spec_parts = [f'{a}' if a == b else f'{a}-{b}' for a, b in ranges]
        self.page_spec.set(','.join(spec_parts))


# ---- スタイル（選択ハイライト枠色など ttk 用） ----
def setup_styles():
    style = ttk.Style()
    # style.configure('Selected.TFrame', borderwidth=4, relief='solid', background=PDFExtractorApp.COLOR_SEL)
    # TFrame は直接 highlightbackground を持たないので、tk.Frame を使っている箇所はそのままで良い。
    # ttk のウィジェットに対する共通スタイル設定などを行う場合はここに記述。
    pass


def main():
    root = tk.Tk()
    setup_styles() # ttk スタイル設定
    app = PDFExtractorApp(root)
    root.mainloop()


if __name__ == '__main__':
    main()