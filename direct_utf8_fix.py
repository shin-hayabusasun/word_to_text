#!/usr/bin/env python
# coding: utf-8

import os
import sys
from pathlib import Path

def fix_txt_file_utf8(txt_path):
    """
    テキストファイルをUTF-8で読み込み、修正して再保存する
    """
    try:
        print(f"ファイルを処理中: {txt_path}")
        
        # UTF-8でファイルを読み込む
        with open(txt_path, 'r', encoding='utf-8', errors='ignore') as f:
            content = f.read()
        
        # ファイル名から自動バックアップの作成
        backup_path = f"{txt_path}.backup"
        with open(backup_path, 'w', encoding='utf-8') as f:
            f.write(content)
        print(f"バックアップを作成しました: {backup_path}")
        
        # 整形された内容をUTF-8で書き込み
        with open(txt_path, 'w', encoding='utf-8') as f:
            f.write(content)
        
        print(f"ファイルをUTF-8で再保存しました: {txt_path}")
        return True
    except Exception as e:
        print(f"処理エラー: {str(e)}")
        return False

def main():
    if len(sys.argv) < 2:
        print("使用方法: python direct_utf8_fix.py <テキストファイルパス>")
        return
    
    file_path = sys.argv[1]
    
    if not os.path.exists(file_path):
        print(f"エラー: 指定されたファイル '{file_path}' が存在しません。")
        return
    
    if not file_path.lower().endswith('.txt'):
        print(f"エラー: 指定されたファイル '{file_path}' はテキストファイル(.txt)ではありません。")
        return
    
    success = fix_txt_file_utf8(file_path)
    
    if success:
        print(f"処理成功: {file_path}")
    else:
        print(f"処理失敗: {file_path}")

if __name__ == "__main__":
    main() 