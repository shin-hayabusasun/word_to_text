#!/usr/bin/env python
# coding: utf-8

import re

def final_clean(input_path, output_path):
    print(f"最終クリーニングを実行中: {input_path}")
    
    with open(input_path, 'r', encoding='utf-8', errors='ignore') as f:
        content = f.read()
    
    # 「社外秘」を検出し、その後の文字化け部分をすべて削除する
    pattern = r'社外秘.*'
    cleaned_content = re.sub(pattern, '社外秘', content, flags=re.DOTALL)
    
    # 出力
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(cleaned_content)
    
    print(f"最終クリーニング完了: {output_path}")

if __name__ == "__main__":
    input_file = "最終クリーニング完了.txt"
    output_file = "最終版.txt"
    final_clean(input_file, output_file) 