#!/usr/bin/env python
# coding: utf-8

import os
import sys
import re
from pathlib import Path

def fix_utf8_and_remove_garbled(txt_path, output_path=None):
    """
    テキストファイルをUTF-8で読み込み、文字化け部分を削除してきれいなテキストのみを抽出する
    
    Args:
        txt_path (str): 処理するテキストファイルのパス
        output_path (str, optional): 出力先のパス。指定がない場合は入力ファイルを上書き
    """
    try:
        print(f"ファイルを処理中: {txt_path}")
        
        # 出力パスが指定されていない場合は入力ファイルを上書き
        if output_path is None:
            output_path = txt_path
        
        # ファイル名から自動バックアップの作成
        backup_path = f"{txt_path}.backup"
        
        # UTF-8でファイルを読み込む
        with open(txt_path, 'r', encoding='utf-8', errors='ignore') as f:
            content = f.read()
        
        # バックアップを作成
        with open(backup_path, 'w', encoding='utf-8') as f:
            f.write(content)
        print(f"バックアップを作成しました: {backup_path}")
        
        # 意味のある行のみを抽出
        lines = content.splitlines()
        cleaned_lines = []
        
        for line in lines:
            # 日本語文字が含まれているか、または意味のある記号や英数字のみの行か確認
            jp_chars = re.findall(r'[ぁ-んァ-ヶ一-龠々〆〜]', line)
            meaningful_symbols = re.findall(r'[・。、：；「」『』（）［］【】◆■□◇※！？]', line)
            
            # 文字化けと思われる記号の連続を含む行を除外
            has_garbled = bool(re.search(r'[龠〆〜耐脀砐源氈鰈璀丄溑蛝鯄瀠鮤]+', line))
            
            # 有意義な行であるかどうかを判断
            if (jp_chars or meaningful_symbols) and len(line.strip()) > 2:
                # 文字化けを含む行は除外
                if not has_garbled:
                    cleaned_lines.append(line)
        
        # 整形された内容をUTF-8で書き込み
        cleaned_content = '\n'.join(cleaned_lines)
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(cleaned_content)
        
        print(f"文字化け部分を除去して再保存しました: {output_path}")
        return True
    except Exception as e:
        print(f"処理エラー: {str(e)}")
        return False

def main():
    if len(sys.argv) < 2:
        print("使用方法: python enhanced_utf8_fix.py <テキストファイルパス> [出力ファイルパス]")
        return
    
    file_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else None
    
    if not os.path.exists(file_path):
        print(f"エラー: 指定されたファイル '{file_path}' が存在しません。")
        return
    
    if not file_path.lower().endswith('.txt'):
        print(f"エラー: 指定されたファイル '{file_path}' はテキストファイル(.txt)ではありません。")
        return
    
    success = fix_utf8_and_remove_garbled(file_path, output_path)
    
    if success:
        print(f"処理成功: {file_path}")
    else:
        print(f"処理失敗: {file_path}")

if __name__ == "__main__":
    main() 