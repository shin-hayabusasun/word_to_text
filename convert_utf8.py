#!/usr/bin/env python
# coding: utf-8

import os
import sys
import re
from pathlib import Path

def convert_doc_to_utf8(doc_path, output_path=None):
    """
    ファイルをUTF-8で変換する簡易スクリプト
    """
    try:
        # 出力パスが指定されていない場合は入力ファイルと同じ場所に.txtを作成
        if output_path is None:
            output_path = str(Path(doc_path).with_suffix('.txt'))
        
        print(f"UTF-8変換モードで処理中: {doc_path}")
        
        # バイナリモードでファイルを開く
        with open(doc_path, 'rb') as f:
            content = f.read()
        
        # UTF-8でデコード
        text = content.decode('utf-8', errors='ignore')
        
        # 日本語文字が含まれているか確認
        jp_chars = re.findall(r'[ぁ-んァ-ヶ一-龠々〆〜]', text)
        jp_ratio = len(jp_chars) / max(len(text), 1)
        
        print(f"UTF-8デコード: テキスト長={len(text)}, 日本語文字数={len(jp_chars)}, 比率={jp_ratio:.2%}")
        
        # 不要なバイナリデータやノイズを除去
        text = re.sub(r'[^\x20-\x7E\u3000-\u30FF\u4E00-\u9FFF\u3040-\u309F\uFF00-\uFF9F\u2000-\u206F\n]+', '', text)
        text = re.sub(r'[\x00-\x1F\x7F]', '', text)  # 制御文字を除去
        
        # XMLタグのような文字列を削除
        text = re.sub(r'<[^>]+>', '', text)
        
        # テキストファイルに書き込む
        with open(output_path, 'w', encoding='utf-8') as out_file:
            out_file.write(text)
        
        print(f"UTF-8による変換完了: {output_path}")
        return output_path
    except Exception as e:
        print(f"変換エラー: {str(e)}")
        return None

def main():
    if len(sys.argv) < 2:
        print("使用方法: python convert_utf8.py <ファイルパス>")
        return
    
    file_path = sys.argv[1]
    
    if not os.path.exists(file_path):
        print(f"エラー: 指定されたファイル '{file_path}' が存在しません。")
        return
    
    output_path = convert_doc_to_utf8(file_path)
    
    if output_path:
        print(f"変換成功: {output_path}")
    else:
        print(f"変換失敗: {file_path}")

if __name__ == "__main__":
    main() 