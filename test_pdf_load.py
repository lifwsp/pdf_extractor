#!/usr/bin/env python3
"""
PDFファイルの読み込みとサムネイル生成をテストするスクリプト
"""
import os
import sys
from pypdf import PdfReader
from pdf2image import convert_from_path

def test_pdf_loading(pdf_path):
    """PDFの読み込みテスト"""
    print(f"=" * 60)
    print(f"PDFファイル読み込みテスト")
    print(f"=" * 60)
    print(f"ファイルパス: {pdf_path}")

    if not os.path.exists(pdf_path):
        print(f"❌ ファイルが存在しません")
        return False

    print(f"ファイルサイズ: {os.path.getsize(pdf_path) / 1024 / 1024:.2f} MB")

    # PyPDF読み込みテスト
    print(f"\n--- PyPDF読み込みテスト ---")
    try:
        reader = PdfReader(pdf_path)
        print(f"✅ PyPDF読み込み成功")
        print(f"   暗号化: {reader.is_encrypted}")
        print(f"   総ページ数: {len(reader.pages)}")

        if reader.metadata:
            print(f"   メタデータ:")
            for key, value in reader.metadata.items():
                print(f"     {key}: {value}")
    except Exception as e:
        print(f"❌ PyPDF読み込み失敗: {type(e).__name__}")
        print(f"   メッセージ: {str(e)}")
        return False

    # pdf2image サムネイル生成テスト
    print(f"\n--- pdf2image サムネイル生成テスト ---")
    try:
        print(f"   最初の2ページのサムネイル生成中...")
        images = convert_from_path(
            pdf_path,
            dpi=72,
            fmt='png',
            thread_count=2,
            first_page=1,
            last_page=min(2, len(reader.pages))
        )
        print(f"✅ サムネイル生成成功: {len(images)}枚")
        for idx, img in enumerate(images, 1):
            print(f"   画像{idx}: {img.size}")
    except Exception as e:
        print(f"❌ サムネイル生成失敗: {type(e).__name__}")
        print(f"   メッセージ: {str(e)}")
        print(f"\n   ヒント: Popplerがインストールされていない可能性があります")
        print(f"   インストールコマンド: sudo apt install poppler-utils")
        return False

    print(f"\n{'=' * 60}")
    print(f"✅ すべてのテスト成功！")
    print(f"{'=' * 60}")
    return True

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("使用方法: python test_pdf_load.py <PDFファイルパス>")
        sys.exit(1)

    pdf_path = sys.argv[1]
    success = test_pdf_loading(pdf_path)
    sys.exit(0 if success else 1)
