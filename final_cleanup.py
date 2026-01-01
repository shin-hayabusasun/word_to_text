#!/usr/bin/env python
# coding: utf-8

import re
import os

def clean_file(input_path, output_path):
    print(f"ファイルを処理中: {input_path}")
    
    # ファイルを読み込む
    with open(input_path, 'r', encoding='utf-8', errors='ignore') as f:
        content = f.read()
    
    # 行ごとに処理
    cleaned_lines = []
    current_paragraph = []
    
    for line in content.splitlines():
        line = line.strip()
        if not line:
            # 空行の場合、段落が溜まっていれば追加
            if current_paragraph:
                cleaned_lines.append(' '.join(current_paragraph))
                current_paragraph = []
            cleaned_lines.append('')  # 空行を保存
            continue
        
        # 文字化けの可能性が高い行は除外
        if re.search(r'[龠〆〜耐脀砐源氈鰈璀丄溑蛝鯄瀠鮤]{3,}', line) or \
           re.search(r'[^ぁ-んァ-ン一-龥a-zA-Z0-9 \t.,;:!?()（）「」『』［］【】・。、：；！？… 　]+', line):
            continue  # 文字化けのある行はスキップ
        
        # 日本語が含まれるか、意味のある英数字や記号が含まれる行のみ保持
        has_japanese = bool(re.search(r'[ぁ-んァ-ヶ一-龠々〆〜]', line))
        has_meaningful_text = bool(re.search(r'[a-zA-Z0-9]{2,}|[・。、：；「」『』（）［］【】◆■□◇※！？]', line))
        
        if (has_japanese or has_meaningful_text) and len(line) > 1:
            # 残すべき行
            current_paragraph.append(line)
    
    # 最後の段落を確認
    if current_paragraph:
        cleaned_lines.append(' '.join(current_paragraph))
    
    # 結果をファイルに書き込む
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(cleaned_lines))
    
    print(f"処理完了: {output_path}")

if __name__ == "__main__":
    input_file = "201908給付金補助資料（手順書）11支払・不払件数の計上方法 (1).txt"
    output_file = "最終クリーニング済.txt"
    clean_file(input_file, output_file) 