# --- PDF Page Extractor - Flet Version ---
# Flet (Flutter-based Python GUI): モダンなUI
# PDF読み書き: pypdf
# PDF → 画像変換: pdf2image （※Poppler 必須）
# 画像処理: Pillow

import os
import re
import threading
import base64
import gc
import shutil
import tempfile
from io import BytesIO
from typing import List, Set

import flet as ft
from pypdf import PdfReader, PdfWriter
from pdf2image import convert_from_path
from PIL import Image

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


def pil_to_base64(image: Image.Image) -> str:
    """PIL ImageをBase64文字列に変換"""
    buffered = BytesIO()
    image.save(buffered, format="PNG")
    img_str = base64.b64encode(buffered.getvalue()).decode()
    return f"data:image/png;base64,{img_str}"


class PDFExtractorApp:
    THUMB_DPI = 72  # 高解像度サムネイル
    THUMB_SIZE = 200  # サムネイル幅
    
    def __init__(self, page: ft.Page):
        self.page = page
        self.setup_page()
        
        # 状態変数
        self.pdf_path = ""
        self.save_path = ""  # 保存先のフルパス
        self.selected_pages: Set[int] = set()  # 0-index
        self.last_selected_page: int = -1  # 最後に選択されたページ（範囲選択用）
        self.reader: PdfReader = None
        self.thumbnails: List[Image.Image] = []
        self.temp_dirs: List[str] = []  # 一時ディレクトリのリスト
        
        # UI コンポーネント
        self.pdf_path_text = ft.Text("PDFファイルが選択されていません", color=ft.Colors.GREY_600)
        self.page_spec_field = ft.TextField(
            label="ページ指定 (例: 1,3-5)",
            width=280,
            border_radius=8,
            on_change=self.on_page_spec_change
        )
        self.save_path_text = ft.Text("PDF読み込み後に自動設定されます", color=ft.Colors.GREY_600)
        self.split_pages_field = ft.TextField(
            label="分割ページ数",
            width=100,
            border_radius=8
        )
        self.progress_bar = ft.ProgressBar(
            width=280,
            visible=False,
            border_radius=8
        )
        self.status_text = ft.Text("待機中", color=ft.Colors.BLUE_700, weight=ft.FontWeight.W_500)
        self.thumbnail_grid = ft.GridView(
            expand=True,
            runs_count=4,
            max_extent=220,
            child_aspect_ratio=0.7,
            spacing=10,
            run_spacing=10,
            padding=ft.padding.all(10)
        )
        
        self.build_ui()
    
    def setup_page(self):
        """ページの基本設定"""
        self.page.title = "PDF Page Extractor"
        self.page.window_width = 1280
        self.page.window_height = 800
        self.page.window_resizable = False
        self.page.theme_mode = ft.ThemeMode.LIGHT
        self.page.bgcolor = ft.Colors.GREY_50
        
        # Material Design 3テーマ
        self.page.theme = ft.Theme(
            color_scheme_seed=ft.Colors.BLUE,
            use_material3=True
        )
    
    def cleanup_resources(self):
        """前のPDFのリソースをクリア"""
        try:
            print(f"DEBUG: リソースクリーンアップ開始")
            
            # 既存のPdfReaderをクリア
            if self.reader:
                self.reader = None
                print(f"DEBUG: PdfReaderクリア完了")
            
            # サムネイル画像をクリア
            if self.thumbnails:
                for img in self.thumbnails:
                    if hasattr(img, 'close'):
                        try:
                            img.close()
                        except:
                            pass
                self.thumbnails.clear()
                print(f"DEBUG: サムネイル画像クリア完了")
            
            # 一時ディレクトリを削除
            for temp_dir in self.temp_dirs:
                try:
                    if os.path.exists(temp_dir):
                        shutil.rmtree(temp_dir)
                        print(f"DEBUG: 一時ディレクトリ削除完了 - {temp_dir}")
                except Exception as e:
                    print(f"DEBUG: 一時ディレクトリ削除失敗 - {temp_dir}: {e}")
            self.temp_dirs.clear()
            
            # 選択状態をクリア
            self.selected_pages.clear()
            self.last_selected_page = -1
            
            # ガベージコレクションを実行（新しいPDFを開く時のみ）
            gc.collect()
            print(f"DEBUG: ガベージコレクション完了")
            
        except Exception as e:
            print(f"DEBUG: リソースクリーンアップエラー - {type(e).__name__}: {e}")
    
    def build_ui(self):
        """UIの構築"""
        # 左サイドバー
        sidebar = ft.Container(
            width=320,
            bgcolor=ft.Colors.WHITE,
            padding=ft.padding.all(20),
            border_radius=ft.border_radius.only(top_right=12, bottom_right=12),
            shadow=ft.BoxShadow(
                spread_radius=1,
                blur_radius=10,
                color=ft.Colors.BLACK12
            ),
            content=ft.Column([
                # タイトル
                ft.Text(
                    "PDF Page Extractor",
                    size=20,
                    weight=ft.FontWeight.BOLD,
                    color=ft.Colors.BLUE_700
                ),
                ft.Divider(height=20),
                
                # PDF選択セクション
                ft.Text("PDFファイル", weight=ft.FontWeight.W_500),
                ft.ElevatedButton(
                    "PDFを選択",
                    icon=ft.Icons.FOLDER_OPEN,
                    on_click=self.select_pdf,
                    style=ft.ButtonStyle(
                        shape=ft.RoundedRectangleBorder(radius=8)
                    )
                ),
                self.pdf_path_text,
                ft.Divider(height=10),
                
                # ページ指定セクション
                ft.Text("ページ抽出", weight=ft.FontWeight.W_500),
                self.page_spec_field,
                ft.Row([
                    ft.ElevatedButton(
                        "全選択",
                        icon=ft.Icons.SELECT_ALL,
                        on_click=self.select_all_pages,
                        style=ft.ButtonStyle(shape=ft.RoundedRectangleBorder(radius=6))
                    ),
                    ft.ElevatedButton(
                        "全解除",
                        icon=ft.Icons.DESELECT,
                        on_click=self.deselect_all_pages,
                        style=ft.ButtonStyle(shape=ft.RoundedRectangleBorder(radius=6))
                    )
                ], spacing=5),
                ft.ElevatedButton(
                    "保存先を変更",
                    icon=ft.Icons.EDIT,
                    on_click=self.select_save,
                    style=ft.ButtonStyle(
                        shape=ft.RoundedRectangleBorder(radius=8)
                    )
                ),
                self.save_path_text,
                ft.FilledButton(
                    "ページ抽出実行",
                    icon=ft.Icons.DOWNLOAD,
                    on_click=self.run_extract,
                    style=ft.ButtonStyle(
                        shape=ft.RoundedRectangleBorder(radius=8)
                    )
                ),
                ft.Divider(height=20),
                
                # P分割セクション
                ft.Text("ページ分割", weight=ft.FontWeight.W_500),
                ft.Row([
                    self.split_pages_field,
                    ft.Text("ページ毎", color=ft.Colors.GREY_600)
                ]),
                ft.FilledButton(
                    "分割実行",
                    icon=ft.Icons.CONTENT_CUT,
                    on_click=self.run_split,
                    style=ft.ButtonStyle(
                        shape=ft.RoundedRectangleBorder(radius=8)
                    )
                ),
                ft.Divider(height=20),
                
                # ステータス
                self.progress_bar,
                self.status_text,
            ], spacing=8)
        )
        
        # 右メインエリア（サムネイル）
        main_area = ft.Container(
            expand=True,
            padding=ft.padding.all(10),
            content=ft.Column([
                ft.Column([
                    ft.Text(
                        "サムネイル一覧",
                        size=18,
                        weight=ft.FontWeight.W_500,
                        color=ft.Colors.BLUE_700
                    ),
                    ft.Text(
                        "💡 クリック: 単一選択 | Shift+クリック: 範囲選択",
                        size=12,
                        color=ft.Colors.GREY_600,
                        italic=True
                    )
                ], spacing=2),
                ft.Container(
                    expand=True,
                    bgcolor=ft.Colors.WHITE,
                    border_radius=12,
                    shadow=ft.BoxShadow(
                        spread_radius=1,
                        blur_radius=10,
                        color=ft.Colors.BLACK12
                    ),
                    content=self.thumbnail_grid
                )
            ], spacing=10)
        )
        
        # メインレイアウト
        self.page.add(
            ft.Row([
                sidebar,
                main_area
            ], spacing=0, expand=True)
        )
    
    def select_pdf(self, e):
        """PDFファイル選択"""
        def file_picker_result(e: ft.FilePickerResultEvent):
            if e.files:
                file_path = e.files[0].path
                print(f"DEBUG: 選択されたファイルパス - {file_path}")
                print(f"DEBUG: ファイル存在確認 - {os.path.exists(file_path)}")
                
                self.pdf_path = file_path
                self.pdf_path_text.value = os.path.basename(file_path)
                self.pdf_path_text.color = ft.Colors.BLACK
                self.clear_thumbnails()
                self.update_status("PDF読み込み中...")
                self.page.update()
                
                # バックグラウンドでPDF読み込み
                threading.Thread(
                    target=self.load_pdf_and_thumbnails,
                    args=(file_path,),
                    daemon=True
                ).start()
        
        file_picker = ft.FilePicker(on_result=file_picker_result)
        self.page.overlay.append(file_picker)
        self.page.update()
        
        file_picker.pick_files(
            allow_multiple=False,
            allowed_extensions=["pdf"]
        )
    
    def select_save(self, e):
        """保存先選択"""
        def file_picker_result(e: ft.FilePickerResultEvent):
            if e.path:
                self.save_path = e.path  # フルパスを保存
                self.save_path_text.value = os.path.basename(e.path)
                self.save_path_text.color = ft.Colors.BLACK
                self.page.update()
        
        file_picker = ft.FilePicker(on_result=file_picker_result)
        self.page.overlay.append(file_picker)
        self.page.update()
        
        # デフォルトファイル名を設定
        if self.save_path:
            default_name = os.path.basename(self.save_path)
        else:
            default_name = "extracted.pdf"
        
        file_picker.save_file(
            dialog_title="保存先を選択",
            file_name=default_name,
            allowed_extensions=["pdf"]
        )
    
    def load_pdf_and_thumbnails(self, path: str):
        """PDFとサムネイルを読み込み"""
        try:
            print(f"DEBUG: PDF読み込み開始 - {path}")
            print(f"DEBUG: ファイル存在確認 - {os.path.exists(path)}")
            
            # 前のPDFのリソースをクリア
            self.cleanup_resources()
            
            # PDF読み込み
            reader = PdfReader(path)
            print(f"DEBUG: PdfReader作成成功")
            if reader.is_encrypted:
                raise RuntimeError('暗号化されたPDFは扱えません')
            
            self.reader = reader
            total_pages = len(reader.pages)
            print(f"DEBUG: 総ページ数 - {total_pages}")
            
            # デフォルト保存先を設定
            self.set_default_save_path(path)
            
            if total_pages == 0:
                print(f"DEBUG: ページ数が0のため処理終了")
                self.update_status('PDF読み込み完了 (0ページ)')
                return
            
            print(f"DEBUG: プログレスバー設定開始")
            # プログレスバー設定
            self.progress_bar.visible = True
            self.progress_bar.value = 0
            self.update_status('サムネイル生成中...')
            self.page.update()
            print(f"DEBUG: プログレスバー設定完了")
            
            print(f"DEBUG: サムネイル生成開始")
            # サムネイル生成（最大100ページ）
            max_thumb_pages = min(total_pages, 100)
            print(f"DEBUG: サムネイル生成対象ページ数 - {max_thumb_pages}")
            
            try:
                images = convert_from_path(
                    path,
                    dpi=self.THUMB_DPI,
                    fmt='png',
                    thread_count=2,
                    first_page=1,
                    last_page=max_thumb_pages
                )
                print(f"DEBUG: convert_from_path成功 - {len(images)}枚の画像")
            except Exception as convert_error:
                print(f"DEBUG: convert_from_path失敗 - {type(convert_error).__name__}: {convert_error}")
                raise
            
            print(f"DEBUG: 画像リサイズ開始")
            thumbnails = []
            for idx, img in enumerate(images):
                try:
                    # サムネイルリサイズ
                    w, h = img.size
                    print(f"DEBUG: 画像{idx+1}サイズ - {w}x{h}")
                    
                    if w > self.THUMB_SIZE:
                        ratio = self.THUMB_SIZE / w
                        new_size = (int(w * ratio), int(h * ratio))
                        img = img.resize(new_size, Image.Resampling.LANCZOS)
                        print(f"DEBUG: 画像{idx+1}リサイズ後 - {new_size}")
                    
                    thumbnails.append(img)
                    
                    # プログレス更新
                    progress = (idx + 1) / max_thumb_pages
                    self.progress_bar.value = progress
                    print(f"DEBUG: 画像{idx+1}処理完了")
                    
                except Exception as img_error:
                    print(f"DEBUG: 画像{idx+1}処理エラー - {type(img_error).__name__}: {img_error}")
                    raise
            
            print(f"DEBUG: サムネイルリスト設定")
            self.thumbnails = thumbnails
            
            print(f"DEBUG: ファイルベースサムネイル表示を開始")
            self.populate_file_based_thumbnail_grid()
            
            print(f"DEBUG: ステータス更新")
            self.update_status(f'PDF読み込み完了 ({total_pages}ページ)')
            print(f"DEBUG: PDF読み込み処理完了")
            
        except Exception as e:
            self.update_status('PDF読み込み失敗')
            print(f"DEBUG: PDF読み込みエラー - {type(e).__name__}: {str(e)}")
            self.show_error('PDFの読み込みに失敗しました', f"{type(e).__name__}: {str(e)}")
            self.reader = None
        finally:
            print(f"DEBUG: finally ブロック開始")
            try:
                self.progress_bar.visible = False
                print(f"DEBUG: プログレスバー非表示設定")
                self.page.update()
                print(f"DEBUG: ページ更新完了 - finally")
            except Exception as finally_error:
                print(f"DEBUG: finallyブロックエラー - {type(finally_error).__name__}: {finally_error}")
            print(f"DEBUG: load_pdf_and_thumbnails関数終了")
    
    def populate_thumbnail_grid(self):
        """サムネイルグリッドを構築"""
        print(f"DEBUG: populate_thumbnail_grid開始")
        try:
            print(f"DEBUG: グリッドクリア")
            self.thumbnail_grid.controls.clear()
            
            print(f"DEBUG: サムネイル数 - {len(self.thumbnails)}")
            for idx, thumbnail in enumerate(self.thumbnails):
                print(f"DEBUG: サムネイル{idx+1}処理開始")
                page_num = idx + 1  # 1-based
                is_selected = idx in self.selected_pages
                
                try:
                    # Base64変換
                    print(f"DEBUG: Base64変換開始")
                    img_base64 = pil_to_base64(thumbnail)
                    print(f"DEBUG: Base64変換完了")
                except Exception as b64_error:
                    print(f"DEBUG: Base64変換エラー - {type(b64_error).__name__}: {b64_error}")
                    raise
                
                try:
                    # サムネイルカード
                    print(f"DEBUG: カード作成開始")
                    card = ft.Container(
                        content=ft.Column([
                            ft.Stack([
                                # サムネイル画像
                                ft.Container(
                                    content=ft.Image(
                                        src=img_base64,
                                        fit=ft.ImageFit.CONTAIN,
                                        border_radius=8
                                    ),
                                    border_radius=8,
                                    bgcolor=ft.Colors.WHITE,
                                    # animate_scale=ft.Animation(200, ft.AnimationCurve.EASE_OUT)
                                ),
                                # チェックマーク（選択時）
                                ft.Container(
                                    content=ft.Icon(
                                        ft.Icons.CHECK_CIRCLE,
                                        color=ft.Colors.WHITE,
                                        size=24
                                    ),
                                    bgcolor=ft.Colors.GREEN,
                                    border_radius=20,
                                    width=28,
                                    height=28,
                                    right=5,
                                    top=5,
                                    visible=is_selected
                                )
                            ], width=self.THUMB_SIZE, height=int(self.THUMB_SIZE * 1.4)),
                            # ページ番号
                            ft.Text(
                                f"ページ {page_num}",
                                size=12,
                                weight=ft.FontWeight.W_500,
                                text_align=ft.TextAlign.CENTER,
                                color=ft.Colors.BLUE_700 if is_selected else ft.Colors.GREY_700
                            )
                        ], spacing=5, horizontal_alignment=ft.CrossAxisAlignment.CENTER),
                        padding=8,
                        border_radius=12,
                        bgcolor=ft.Colors.BLUE_50 if is_selected else ft.Colors.WHITE,
                        border=ft.border.all(
                            2, ft.Colors.BLUE_400 if is_selected else ft.Colors.GREY_300
                        ),
                        shadow=ft.BoxShadow(
                            spread_radius=1,
                            blur_radius=5,
                            color=ft.Colors.BLACK12
                        ) if is_selected else None,
                        on_click=lambda e, page=page_num: self.toggle_page(page),
                        # animate=ft.Animation(200, ft.AnimationCurve.EASE_OUT),
                        scale=1.05 if is_selected else 1.0
                    )
                    print(f"DEBUG: カード作成完了")
                    
                    print(f"DEBUG: グリッドに追加")
                    self.thumbnail_grid.controls.append(card)
                    print(f"DEBUG: サムネイル{idx+1}処理完了")
                    
                except Exception as card_error:
                    print(f"DEBUG: カード作成エラー - {type(card_error).__name__}: {card_error}")
                    raise
            
            print(f"DEBUG: ページ更新開始")
            try:
                self.page.update()
                print(f"DEBUG: ページ更新完了")
            except Exception as update_error:
                print(f"DEBUG: ページ更新エラー - {type(update_error).__name__}: {update_error}")
                # 更新エラーでも処理は継続
            
        except Exception as grid_error:
            print(f"DEBUG: populate_thumbnail_gridエラー - {type(grid_error).__name__}: {grid_error}")
            raise
    
    def populate_simple_thumbnail_grid(self):
        """簡易サムネイルグリッドを構築（画像なしでページ番号のみ）"""
        print(f"DEBUG: populate_simple_thumbnail_grid開始")
        try:
            print(f"DEBUG: グリッドクリア")
            self.thumbnail_grid.controls.clear()
            
            print(f"DEBUG: サムネイル数 - {len(self.thumbnails)}")
            for idx in range(len(self.thumbnails)):
                page_num = idx + 1  # 1-based
                is_selected = idx in self.selected_pages
                
                print(f"DEBUG: 簡易カード{idx+1}作成開始")
                
                # 画像なしの簡易カード
                card = ft.Container(
                    content=ft.Column([
                        # ページ番号のみ表示（画像の代わり）
                        ft.Container(
                            content=ft.Text(
                                f"ページ\n{page_num}",
                                size=16,
                                weight=ft.FontWeight.BOLD,
                                text_align=ft.TextAlign.CENTER,
                                color=ft.Colors.BLUE_700
                            ),
                            width=self.THUMB_SIZE,
                            height=int(self.THUMB_SIZE * 1.4),
                            bgcolor=ft.Colors.GREY_100,
                            border_radius=8,
                            alignment=ft.alignment.center
                        ),
                        # チェックマーク（選択時）
                        ft.Text(
                            "✓ 選択済み" if is_selected else "クリックで選択",
                            size=10,
                            color=ft.Colors.GREEN if is_selected else ft.Colors.GREY_600,
                            text_align=ft.TextAlign.CENTER
                        )
                    ], spacing=5, horizontal_alignment=ft.CrossAxisAlignment.CENTER),
                    padding=8,
                    border_radius=12,
                    bgcolor=ft.Colors.BLUE_50 if is_selected else ft.Colors.WHITE,
                    border=ft.border.all(
                        2, ft.Colors.BLUE_400 if is_selected else ft.Colors.GREY_300
                    ),
                    on_click=lambda e, page=page_num: self.toggle_page(page),
                    scale=1.05 if is_selected else 1.0
                )
                
                print(f"DEBUG: 簡易カード{idx+1}作成完了")
                self.thumbnail_grid.controls.append(card)
            
            print(f"DEBUG: 簡易ページ更新開始")
            self.page.update()
            print(f"DEBUG: 簡易ページ更新完了")
            
        except Exception as simple_error:
            print(f"DEBUG: populate_simple_thumbnail_gridエラー - {type(simple_error).__name__}: {simple_error}")
            raise
    
    def populate_real_thumbnail_grid(self):
        """実際のサムネイルグリッドを構築（段階的実装）"""
        print(f"DEBUG: populate_real_thumbnail_grid開始")
        try:
            print(f"DEBUG: グリッドクリア")
            self.thumbnail_grid.controls.clear()
            
            print(f"DEBUG: サムネイル数 - {len(self.thumbnails)}")
            
            # 最初の3つだけ画像表示を試す
            max_display = min(3, len(self.thumbnails))
            print(f"DEBUG: 表示する画像数 - {max_display}")
            
            for idx in range(len(self.thumbnails)):
                page_num = idx + 1  # 1-based
                is_selected = idx in self.selected_pages
                
                print(f"DEBUG: カード{idx+1}作成開始")
                
                if idx < max_display:
                    # 画像付きカード（最初の3つのみ）
                    try:
                        print(f"DEBUG: Base64変換開始 - 画像{idx+1}")
                        img_base64 = pil_to_base64(self.thumbnails[idx])
                        print(f"DEBUG: Base64変換完了 - 画像{idx+1}")
                        
                        card_content = ft.Column([
                            ft.Container(
                                content=ft.Image(
                                    src=img_base64,
                                    fit=ft.ImageFit.CONTAIN,
                                    border_radius=8
                                ),
                                width=self.THUMB_SIZE,
                                height=int(self.THUMB_SIZE * 1.4),
                                bgcolor=ft.Colors.WHITE,
                                border_radius=8
                            ),
                            ft.Text(
                                f"ページ {page_num}",
                                size=12,
                                weight=ft.FontWeight.W_500,
                                text_align=ft.TextAlign.CENTER,
                                color=ft.Colors.BLUE_700 if is_selected else ft.Colors.GREY_700
                            )
                        ], spacing=5, horizontal_alignment=ft.CrossAxisAlignment.CENTER)
                        
                        print(f"DEBUG: 画像カード{idx+1}作成完了")
                        
                    except Exception as img_error:
                        print(f"DEBUG: 画像カード{idx+1}作成失敗 - {type(img_error).__name__}: {img_error}")
                        # 失敗した場合はテキストカードにフォールバック
                        card_content = ft.Column([
                            ft.Container(
                                content=ft.Text(
                                    f"ページ\n{page_num}\n(画像エラー)",
                                    size=14,
                                    weight=ft.FontWeight.BOLD,
                                    text_align=ft.TextAlign.CENTER,
                                    color=ft.Colors.RED
                                ),
                                width=self.THUMB_SIZE,
                                height=int(self.THUMB_SIZE * 1.4),
                                bgcolor=ft.Colors.GREY_100,
                                border_radius=8,
                                alignment=ft.alignment.center
                            )
                        ], spacing=5, horizontal_alignment=ft.CrossAxisAlignment.CENTER)
                else:
                    # テキストのみカード（4つ目以降）
                    card_content = ft.Column([
                        ft.Container(
                            content=ft.Text(
                                f"ページ\n{page_num}",
                                size=16,
                                weight=ft.FontWeight.BOLD,
                                text_align=ft.TextAlign.CENTER,
                                color=ft.Colors.BLUE_700
                            ),
                            width=self.THUMB_SIZE,
                            height=int(self.THUMB_SIZE * 1.4),
                            bgcolor=ft.Colors.GREY_100,
                            border_radius=8,
                            alignment=ft.alignment.center
                        )
                    ], spacing=5, horizontal_alignment=ft.CrossAxisAlignment.CENTER)
                
                # カードを作成
                card = ft.Container(
                    content=card_content,
                    padding=8,
                    border_radius=12,
                    bgcolor=ft.Colors.BLUE_50 if is_selected else ft.Colors.WHITE,
                    border=ft.border.all(
                        2, ft.Colors.BLUE_400 if is_selected else ft.Colors.GREY_300
                    ),
                    on_click=lambda e, page_num=page_num: self.handle_page_click(e, page_num),
                    scale=1.05 if is_selected else 1.0
                )
                
                print(f"DEBUG: カード{idx+1}グリッドに追加")
                self.thumbnail_grid.controls.append(card)
            
            print(f"DEBUG: 実際のページ更新開始")
            self.page.update()
            print(f"DEBUG: 実際のページ更新完了")
            
        except Exception as real_error:
            print(f"DEBUG: populate_real_thumbnail_gridエラー - {type(real_error).__name__}: {real_error}")
            # エラーが発生した場合は簡易版にフォールバック
            print(f"DEBUG: 簡易版にフォールバック")
            self.populate_simple_thumbnail_grid()
    
    def populate_file_based_thumbnail_grid(self):
        """ファイルベースのサムネイルグリッドを構築"""
        print(f"DEBUG: populate_file_based_thumbnail_grid開始")
        try:
            print(f"DEBUG: グリッドクリア")
            self.thumbnail_grid.controls.clear()
            
            print(f"DEBUG: サムネイル数 - {len(self.thumbnails)}")
            
            # 一時ディレクトリの準備
            temp_dir = tempfile.mkdtemp()
            self.temp_dirs.append(temp_dir)  # リストに追加して後で削除できるように
            print(f"DEBUG: 一時ディレクトリ - {temp_dir}")
            
            # 全てのサムネイルを画像表示
            max_display = len(self.thumbnails)
            print(f"DEBUG: 表示する画像数 - {max_display}")
            
            thumbnail_files = []
            for idx in range(max_display):
                try:
                    print(f"DEBUG: 画像{idx+1}ファイル保存開始")
                    temp_path = os.path.join(temp_dir, f"thumb_{idx+1}.png")
                    self.thumbnails[idx].save(temp_path, "PNG")
                    thumbnail_files.append(temp_path)
                    print(f"DEBUG: 画像{idx+1}ファイル保存完了 - {temp_path}")
                except Exception as save_error:
                    print(f"DEBUG: 画像{idx+1}ファイル保存失敗 - {type(save_error).__name__}: {save_error}")
                    thumbnail_files.append(None)
            
            # グリッドに追加
            for idx in range(len(self.thumbnails)):
                page_num = idx + 1  # 1-based
                is_selected = idx in self.selected_pages
                
                print(f"DEBUG: カード{idx+1}作成開始")
                
                if idx < len(thumbnail_files) and thumbnail_files[idx] is not None:
                    # ファイルベース画像カード
                    try:
                        print(f"DEBUG: ファイルベース画像カード{idx+1}作成開始")
                        
                        card_content = ft.Column([
                            ft.Container(
                                content=ft.Image(
                                    src=thumbnail_files[idx],
                                    fit=ft.ImageFit.CONTAIN,
                                    border_radius=8
                                ),
                                width=self.THUMB_SIZE,
                                height=int(self.THUMB_SIZE * 1.4),
                                bgcolor=ft.Colors.WHITE,
                                border_radius=8
                            ),
                            ft.Text(
                                f"ページ {page_num}",
                                size=12,
                                weight=ft.FontWeight.W_500,
                                text_align=ft.TextAlign.CENTER,
                                color=ft.Colors.BLUE_700 if is_selected else ft.Colors.GREY_700
                            )
                        ], spacing=5, horizontal_alignment=ft.CrossAxisAlignment.CENTER)
                        
                        print(f"DEBUG: ファイルベース画像カード{idx+1}作成完了")
                        
                    except Exception as file_img_error:
                        print(f"DEBUG: ファイルベース画像カード{idx+1}作成失敗 - {type(file_img_error).__name__}: {file_img_error}")
                        # 失敗した場合はテキストカードにフォールバック
                        card_content = ft.Column([
                            ft.Container(
                                content=ft.Text(
                                    f"ページ\n{page_num}\n(画像エラー)",
                                    size=14,
                                    weight=ft.FontWeight.BOLD,
                                    text_align=ft.TextAlign.CENTER,
                                    color=ft.Colors.RED
                                ),
                                width=self.THUMB_SIZE,
                                height=int(self.THUMB_SIZE * 1.4),
                                bgcolor=ft.Colors.GREY_100,
                                border_radius=8,
                                alignment=ft.alignment.center
                            )
                        ], spacing=5, horizontal_alignment=ft.CrossAxisAlignment.CENTER)
                else:
                    # テキストのみカード
                    card_content = ft.Column([
                        ft.Container(
                            content=ft.Text(
                                f"ページ\n{page_num}",
                                size=16,
                                weight=ft.FontWeight.BOLD,
                                text_align=ft.TextAlign.CENTER,
                                color=ft.Colors.BLUE_700
                            ),
                            width=self.THUMB_SIZE,
                            height=int(self.THUMB_SIZE * 1.4),
                            bgcolor=ft.Colors.GREY_100,
                            border_radius=8,
                            alignment=ft.alignment.center
                        )
                    ], spacing=5, horizontal_alignment=ft.CrossAxisAlignment.CENTER)
                
                # カードを作成
                card = ft.Container(
                    content=card_content,
                    padding=8,
                    border_radius=12,
                    bgcolor=ft.Colors.BLUE_50 if is_selected else ft.Colors.WHITE,
                    border=ft.border.all(
                        2, ft.Colors.BLUE_400 if is_selected else ft.Colors.GREY_300
                    ),
                    on_click=lambda e, page_num=page_num: self.toggle_page(page_num),
                    scale=1.05 if is_selected else 1.0
                )
                
                print(f"DEBUG: カード{idx+1}グリッドに追加")
                self.thumbnail_grid.controls.append(card)
            
            print(f"DEBUG: ファイルベースページ更新開始")
            self.page.update()
            print(f"DEBUG: ファイルベースページ更新完了")
            
        except Exception as file_error:
            print(f"DEBUG: populate_file_based_thumbnail_gridエラー - {type(file_error).__name__}: {file_error}")
            # エラーが発生した場合は簡易版にフォールバック
            print(f"DEBUG: 簡易版にフォールバック")
            self.populate_simple_thumbnail_grid()
    
    def handle_page_click(self, e, page_number: int):
        """ページクリック処理（Shift対応）"""
        print(f"DEBUG: handle_page_click開始 - ページ{page_number}")
        try:
            # Shiftキーが押されているかチェック
            shift_pressed = hasattr(e, 'shift') and e.shift
            print(f"DEBUG: Shiftキー状態 - {shift_pressed}")
            
            if shift_pressed and self.last_selected_page >= 0:
                # Shift+クリック：範囲選択
                self.select_page_range(self.last_selected_page, page_number)
            else:
                # 通常のクリック：単一選択/解除
                self.toggle_single_page(page_number)
            
            # 最後に選択されたページを更新
            self.last_selected_page = page_number
            
            self.update_page_spec_field()
            self.populate_file_based_thumbnail_grid()
            
        except Exception as click_error:
            print(f"DEBUG: handle_page_clickエラー - {type(click_error).__name__}: {click_error}")
            raise
    
    def toggle_single_page(self, page_number: int):
        """単一ページの選択/解除をトグル"""
        print(f"DEBUG: toggle_single_page開始 - ページ{page_number}")
        idx = page_number - 1  # 0-based
        
        if idx in self.selected_pages:
            self.selected_pages.remove(idx)
            print(f"DEBUG: ページ{page_number}を選択解除")
        else:
            self.selected_pages.add(idx)
            print(f"DEBUG: ページ{page_number}を選択")
    
    def select_page_range(self, start_page: int, end_page: int):
        """ページ範囲選択"""
        print(f"DEBUG: select_page_range開始 - {start_page}から{end_page}")
        
        # 範囲を正規化（小さい値から大きい値へ）
        min_page = min(start_page, end_page)
        max_page = max(start_page, end_page)
        
        print(f"DEBUG: 正規化された範囲 - {min_page}から{max_page}")
        
        # 範囲内のすべてのページを選択
        for page_num in range(min_page, max_page + 1):
            idx = page_num - 1  # 0-based
            if 0 <= idx < len(self.thumbnails):  # 有効な範囲内かチェック
                self.selected_pages.add(idx)
                print(f"DEBUG: ページ{page_num}を範囲選択で追加")
    
    def toggle_page(self, page_number: int):
        """ページ選択/解除をトグル（後方互換性のため残す）"""
        self.toggle_single_page(page_number)
        self.update_page_spec_field()
        self.populate_file_based_thumbnail_grid()
    
    def select_all_pages(self, e):
        """全ページを選択"""
        print(f"DEBUG: select_all_pages開始")
        if self.thumbnails:
            self.selected_pages = set(range(len(self.thumbnails)))
            self.last_selected_page = len(self.thumbnails)
            print(f"DEBUG: 全{len(self.thumbnails)}ページを選択")
            self.update_page_spec_field()
            self.populate_file_based_thumbnail_grid()
    
    def deselect_all_pages(self, e):
        """全ページ選択を解除"""
        print(f"DEBUG: deselect_all_pages開始")
        self.selected_pages.clear()
        self.last_selected_page = -1
        print(f"DEBUG: 全ページの選択を解除")
        self.update_page_spec_field()
        self.populate_file_based_thumbnail_grid()
    
    def on_page_spec_change(self, e):
        """ページ指定フィールドの変更時に選択状態を同期"""
        if not self.thumbnails or not e.control.value.strip():
            return
        
        try:
            # テキストから選択ページを解析
            pages = parse_pages(e.control.value, len(self.thumbnails))
            # 選択状態を更新（0-basedに変換）
            self.selected_pages = set(pages)
            # サムネイル表示を更新（再描画を避けて効率化）
            # self.populate_file_based_thumbnail_grid()
        except Exception:
            # 無効な入力の場合は何もしない
            pass
    
    def set_default_save_path(self, pdf_path: str):
        """PDFファイルパスからデフォルト保存先を設定"""
        try:
            # ファイル名と拡張子を分離
            dir_name = os.path.dirname(pdf_path)
            base_name = os.path.splitext(os.path.basename(pdf_path))[0]
            
            # デフォルト保存先パスを生成（_extr.pdf サフィックス）
            default_save_path = os.path.join(dir_name, f"{base_name}_extr.pdf")
            
            # 保存先を設定
            self.save_path = default_save_path
            self.save_path_text.value = f"{base_name}_extr.pdf"
            self.save_path_text.color = ft.Colors.BLACK
            
            print(f"DEBUG: デフォルト保存先設定 - {default_save_path}")
            self.page.update()
            
        except Exception as e:
            print(f"DEBUG: デフォルト保存先設定エラー - {type(e).__name__}: {e}")
    
    def update_page_spec_field(self):
        """選択ページをテキストフィールドに反映"""
        print(f"DEBUG: update_page_spec_field開始")
        try:
            if not self.selected_pages:
                print(f"DEBUG: 選択ページなし")
                self.page_spec_field.value = ""
                self.page.update()
                return
            
            print(f"DEBUG: 選択ページ数 - {len(self.selected_pages)}")
            # 1-basedに変換してソート
            pages_1based = sorted([p + 1 for p in self.selected_pages])
            print(f"DEBUG: pages_1based - {pages_1based}")
            
            # 連続する範囲をまとめる
            ranges = []
            start = prev = pages_1based[0]
            
            for page in pages_1based[1:]:
                if page == prev + 1:
                    prev = page
                else:
                    ranges.append((start, prev))
                    start = prev = page
            ranges.append((start, prev))
            
            print(f"DEBUG: ranges - {ranges}")
            # 文字列に変換
            spec_parts = [f'{a}' if a == b else f'{a}-{b}' for a, b in ranges]
            spec_text = ','.join(spec_parts)
            print(f"DEBUG: spec_text - {spec_text}")
            
            self.page_spec_field.value = spec_text
            print(f"DEBUG: page.update開始")
            self.page.update()
            print(f"DEBUG: page.update完了")
            
        except Exception as update_error:
            print(f"DEBUG: update_page_spec_fieldエラー - {type(update_error).__name__}: {update_error}")
            raise
    
    def clear_thumbnails(self):
        """サムネイルをクリア"""
        self.thumbnail_grid.controls.clear()
        self.thumbnails.clear()
        self.selected_pages.clear()
        self.last_selected_page = -1
        self.page_spec_field.value = ""
        
        # 保存先もリセット
        self.save_path = ""
        self.save_path_text.value = "PDF読み込み後に自動設定されます"
        self.save_path_text.color = ft.Colors.GREY_600
        
        self.page.update()
    
    def update_status(self, message: str):
        """ステータス更新"""
        self.status_text.value = message
        self.page.update()
    
    def show_error(self, title: str, message: str):
        """エラーダイアログ表示"""
        def close_dialog(e):
            dialog.open = False
            self.page.update()
        
        dialog = ft.AlertDialog(
            title=ft.Text(title),
            content=ft.Text(message),
            actions=[
                ft.TextButton("OK", on_click=close_dialog)
            ]
        )
        
        self.page.dialog = dialog
        dialog.open = True
        self.page.update()
    
    def show_info(self, title: str, message: str):
        """情報ダイアログ表示"""
        def close_dialog(e):
            dialog.open = False
            self.page.update()
        
        dialog = ft.AlertDialog(
            title=ft.Text(title),
            content=ft.Text(message),
            actions=[
                ft.TextButton("OK", on_click=close_dialog)
            ]
        )
        
        self.page.dialog = dialog
        dialog.open = True
        self.page.update()
    
    def run_extract(self, e):
        """ページ抽出実行"""
        if not self.reader:
            self.show_error('エラー', '先にPDFを読み込んでください')
            return
        
        try:
            pages = parse_pages(self.page_spec_field.value, len(self.reader.pages))
            if not pages:
                raise ValueError('抽出ページが指定されていません')
            
            if not self.save_path:
                raise ValueError('保存先が指定されていません')
            
            # バックグラウンドで実行
            threading.Thread(
                target=self.extract_pages,
                args=(pages,),
                daemon=True
            ).start()
            
        except Exception as ex:
            self.show_error('エラー', str(ex))
    
    def extract_pages(self, pages: List[int]):
        """ページ抽出処理"""
        try:
            self.progress_bar.visible = True
            self.progress_bar.value = 0
            self.update_status('ページ抽出中...')
            
            writer = PdfWriter()
            for i, page_idx in enumerate(pages):
                writer.add_page(self.reader.pages[page_idx])
                progress = (i + 1) / len(pages)
                self.progress_bar.value = progress
                self.page.update()
            
            if self.reader.metadata:
                writer.add_metadata(self.reader.metadata)
            
            # ファイル保存
            with open(self.save_path, 'wb') as f:
                writer.write(f)
            
            self.update_status(f'抽出完了 ({len(pages)}ページ)')
            self.show_info('完了', 'PDFを保存しました')
            
        except Exception as e:
            self.update_status('抽出失敗')
            self.show_error('エラー', str(e))
        finally:
            self.progress_bar.visible = False
            self.page.update()
    
    def run_split(self, e):
        """P分割実行"""
        if not self.reader:
            self.show_error('エラー', '先にPDFを読み込んでください')
            return
        
        try:
            p_str = self.split_pages_field.value.strip()
            if not p_str.isdigit() or int(p_str) <= 0:
                raise ValueError('分割ページ数Pは正の整数で指定してください')
            
            p_val = int(p_str)
            
            # バックグラウンドで実行
            threading.Thread(
                target=self.split_pages,
                args=(p_val,),
                daemon=True
            ).start()
            
        except Exception as ex:
            self.show_error('エラー', str(ex))
    
    def split_pages(self, p_val: int):
        """ページ分割処理"""
        try:
            total_pages = len(self.reader.pages)
            num_output_files = (total_pages + p_val - 1) // p_val
            
            self.progress_bar.visible = True
            self.progress_bar.value = 0
            self.update_status(f'{p_val}ページ毎に分割中...')
            
            # 出力ディレクトリを決定
            if self.save_path:
                output_dir = os.path.dirname(self.save_path)
                base_name = os.path.splitext(os.path.basename(self.save_path))[0]
            else:
                output_dir = os.path.dirname(self.pdf_path)
                base_name = os.path.splitext(os.path.basename(self.pdf_path))[0]
            
            generated_files = 0
            for i in range(num_output_files):
                writer = PdfWriter()
                start_idx = i * p_val
                end_idx = min((i + 1) * p_val, total_pages)
                
                for page_idx in range(start_idx, end_idx):
                    writer.add_page(self.reader.pages[page_idx])
                
                if self.reader.metadata:
                    writer.add_metadata(self.reader.metadata)
                
                # ファイル名生成
                start_page = start_idx + 1
                end_page = end_idx
                
                if p_val == 1:
                    filename = f"{base_name}_page_{start_page}.pdf"
                else:
                    filename = f"{base_name}_pages_{start_page}-{end_page}.pdf"
                
                filepath = os.path.join(output_dir, filename)
                
                # 上書き確認
                if os.path.exists(filepath):
                    # 簡略化：自動で上書き
                    pass
                
                with open(filepath, 'wb') as f:
                    writer.write(f)
                
                generated_files += 1
                progress = (i + 1) / num_output_files
                self.progress_bar.value = progress
                self.page.update()
            
            self.update_status(f'分割完了 ({generated_files}個のファイルを出力)')
            self.show_info('完了', f'{generated_files}個のPDFファイルを\n{output_dir}\nに保存しました')
            
        except Exception as e:
            self.update_status('分割失敗')
            self.show_error('エラー', str(e))
        finally:
            self.progress_bar.visible = False
            self.page.update()


def main(page: ft.Page):
    """メイン関数"""
    app = PDFExtractorApp(page)


if __name__ == '__main__':
    ft.app(target=main)