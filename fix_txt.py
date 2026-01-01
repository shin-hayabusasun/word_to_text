#!/usr/bin/env python
# coding: utf-8

import re
import os

def fix_text_file(input_path, output_path):
    # ファイルを読み込む
    with open(input_path, 'r', encoding='utf-8', errors='ignore') as f:
        content = f.read()
    
    # 行ごとに処理
    cleaned_lines = []
    for line in content.splitlines():
        # 文字化けの検出
        if re.search(r'[龠〆〜耐脀砐源氈鰈璀丄溑蛝鯄瀠鮤]{2,}', line):
            continue  # 文字化けのある行はスキップ
        
        # 日本語が含まれるか、意味のある記号が含まれる行のみ保持
        if re.search(r'[ぁ-んァ-ヶ一-龠々〆〜]|[・。、：；「」『』（）［］【】◆■□◇※！？]', line) and len(line.strip()) > 2:
            cleaned_lines.append(line)
    
    # 結果をファイルに書き込む
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(cleaned_lines))
    
    print(f"処理完了: {output_path}")

if __name__ == "__main__":
    fix_text_file(
        "201908給付金補助資料（手順書）11支払・不払件数の計上方法 (1).txt",
        "cleaned.txt"
    ) 