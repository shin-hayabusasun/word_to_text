#!/usr/bin/env python
# coding: utf-8

import os
import sys
import re
from pathlib import Path

def clean_text(input_path, output_path=None, encoding='utf-8'):
    """
    テキストファイルをクリーニングし、読みやすいテキストだけを抽出する
    """
    # 出力パスが指定されていない場合は入力ファイルに _cleaned を付加
    if output_path is None:
        output_path = str(Path(input_path).with_suffix('')) + "_cleaned.txt"
    
    print(f"クリーニング中: {input_path}")
    print(f"出力先: {output_path}")
    
    try:
        # ファイルを読み込む
        with open(input_path, 'r', encoding=encoding, errors='ignore') as f:
            content = f.read()
        
        # ステップ1: 明らかなバイナリデータやXMLを削除
        cleaned_text = re.sub(r'<\?xml.*?>', '', content, flags=re.DOTALL)
        cleaned_text = re.sub(r'</?[a-zA-Z0-9:]+.*?>', '', cleaned_text, flags=re.DOTALL)
        
        # ステップ2: 制御文字やバイナリゴミを削除
        cleaned_text = re.sub(r'[^\x20-\x7E\u3000-\u30FF\u4E00-\u9FFF\u3040-\u309F\uFF00-\uFF9F\u2000-\u206F\n]+', ' ', cleaned_text)
        
        # ステップ3: 日本語文字と英数字、句読点のみの行を抽出
        lines = cleaned_text.split('\n')
        meaningful_lines = []
        
        for line in lines:
            # 基本的なクリーンアップ
            line = line.strip()
            
            # 空行をスキップ
            if not line:
                continue
            
            # 日本語文字または英数字を含む行のみを保持
            if re.search(r'[ぁ-んァ-ヶ一-龠々〆〜]', line) or re.search(r'[a-zA-Z0-9]{3,}', line):
                # さらに不要な記号を削除
                line = re.sub(r'[\x00-\x1F\x7F]', '', line)  # 制御文字を削除
                
                # 行に十分な文字数がある場合のみ追加
                if len(line) > 3:
                    meaningful_lines.append(line)
        
        # 重複行を削除
        unique_lines = []
        seen = set()
        
        for line in meaningful_lines:
            if line not in seen:
                seen.add(line)
                unique_lines.append(line)
        
        # クリーニング済みテキストを書き込む
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(unique_lines))
        
        print(f"クリーニング完了: {output_path}")
        print(f"元の行数: {len(lines)}, クリーニング後の行数: {len(unique_lines)}")
        return True
    
    except Exception as e:
        print(f"エラー: {str(e)}")
        return False

def main():
    if len(sys.argv) < 2:
        print("使用方法: python cleanup_text.py <テキストファイル> [--encoding=ENCODING] [--output=OUTPUT_PATH]")
        print("  --encoding=ENCODING: 入力ファイルのエンコーディング (デフォルト: utf-8)")
        print("  --output=OUTPUT_PATH: 出力先ファイルパス (デフォルト: 入力ファイル名_cleaned.txt)")
        return
    
    input_path = sys.argv[1]
    encoding = 'utf-8'
    output_path = None
    
    # コマンドライン引数を解析
    for arg in sys.argv[2:]:
        if arg.startswith("--encoding="):
            encoding = arg.split("=")[1]
        elif arg.startswith("--output="):
            output_path = arg.split("=")[1]
    
    if not os.path.exists(input_path):
        print(f"エラー: 指定されたファイル '{input_path}' が存在しません。")
        return
    
    # ファイルをクリーニング
    if clean_text(input_path, output_path, encoding):
        print(f"クリーニングが完了しました。")
    else:
        print(f"クリーニングに失敗しました。")

if __name__ == "__main__":
    main() 