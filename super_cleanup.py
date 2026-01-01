#!/usr/bin/env python
# coding: utf-8

import re
import os

def super_clean_file(input_path, output_path):
    print(f"高度なクリーニングを実行中: {input_path}")
    
    # ファイルを読み込む
    with open(input_path, 'r', encoding='utf-8', errors='ignore') as f:
        content = f.read()
    
    # 段落単位で処理
    paragraphs = []
    current_para = []
    
    # 文字化けチェック関数
    def is_garbled(text):
        # 文字化けの特徴的なパターン
        if re.search(r'[龠〆〜耐脀砐源氈鰈璀丄溑蛝鯄瀠鮤]{2,}', text):
            return True
        # XMLや特殊記号の連続パターン
        if re.search(r'[<>/\\{}[\]|@#$%^&*=+`~]{4,}', text):
            return True
        # 意味不明な文字の連続
        if re.search(r'[^ぁ-んァ-ン一-龥a-zA-Z0-9 \t.,;:!?()（）「」『』［］【】・。、：；！？…　]{4,}', text):
            return True
        return False
    
    # 日本語らしさチェック関数
    def has_japanese_content(text):
        # 日本語文字が含まれているか
        if re.search(r'[ぁ-んァ-ヶ一-龠々〆〜]', text):
            return True
        # 意味のある記号や数字だけでも有効とする
        if re.search(r'[a-zA-Z0-9]{2,}|[・。、：；「」『』（）［］【】◆■□◇※！？]', text) and len(text) > 3:
            return True
        return False
    
    # 行ごとに処理
    lines = content.splitlines()
    for line in lines:
        line = line.strip()
        
        # 空行処理
        if not line:
            if current_para:
                paragraph_text = ' '.join(current_para)
                # 文字化けがなく、日本語コンテンツがある段落のみ保持
                if not is_garbled(paragraph_text) and has_japanese_content(paragraph_text):
                    paragraphs.append(paragraph_text)
                current_para = []
            continue
        
        # 文字化けをチェック
        if not is_garbled(line) and has_japanese_content(line):
            current_para.append(line)
    
    # 最後の段落を処理
    if current_para:
        paragraph_text = ' '.join(current_para)
        if not is_garbled(paragraph_text) and has_japanese_content(paragraph_text):
            paragraphs.append(paragraph_text)
    
    # 結果をファイルに書き込む
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write('\n\n'.join(paragraphs))
    
    print(f"クリーニング完了: {output_path}")
    print(f"元のファイル行数: {len(lines)}, クリーニング後の段落数: {len(paragraphs)}")

if __name__ == "__main__":
    input_file = "201908給付金補助資料（手順書）11支払・不払件数の計上方法 (1).txt"
    output_file = "最終クリーニング完了.txt"
    super_clean_file(input_file, output_file) 