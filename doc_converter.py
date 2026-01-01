#!/usr/bin/env python
# coding: utf-8

import os
import sys
import win32com.client
import tempfile
import shutil
import docx2txt
from pathlib import Path

def convert_doc_file(doc_path, output_path=None, encoding='utf-8'):
    """
    .docファイルをテキストファイルに変換する
    """
    try:
        # 出力パスが指定されていない場合は入力ファイルと同じ場所に.txtを作成
        if output_path is None:
            output_path = str(Path(doc_path).with_suffix('.txt'))
        
        print(f"変換中: {doc_path}")
        print(f"出力先: {output_path}")
        
        # 一時ディレクトリを作成
        temp_dir = tempfile.mkdtemp()
        temp_docx = os.path.join(temp_dir, "temp.docx")
        
        try:
            # Word COMを使用して.docを.docxに変換
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            
            try:
                print("Word COMでdocxに変換中...")
                doc = word.Documents.Open(os.path.abspath(doc_path))
                doc.SaveAs2(os.path.abspath(temp_docx), FileFormat=16)  # 16 = .docx形式
                doc.Close(SaveChanges=False)
                
                print("docxからテキストを抽出中...")
                text = docx2txt.process(temp_docx)
                
                with open(output_path, 'w', encoding=encoding) as f:
                    f.write(text)
                
                print(f"変換成功: {output_path}")
                return True
            except Exception as e:
                print(f"Word COMでの変換に失敗: {str(e)}")
                return False
            finally:
                word.Quit()
        except Exception as e:
            print(f"Word COMの初期化に失敗: {str(e)}")
            return False
        finally:
            try:
                shutil.rmtree(temp_dir)
            except:
                pass
    
    except Exception as e:
        print(f"変換エラー: {str(e)}")
        return False

def main():
    if len(sys.argv) < 2:
        print("使用方法: python doc_converter.py <docファイル> [--encoding=ENCODING]")
        print("  --encoding=ENCODING: 出力エンコーディング (デフォルト: utf-8, 例: --encoding=shift-jis)")
        return
    
    doc_path = sys.argv[1]
    encoding = 'utf-8'
    
    # コマンドライン引数を解析
    for arg in sys.argv[2:]:
        if arg.startswith("--encoding="):
            encoding = arg.split("=")[1]
    
    if not os.path.exists(doc_path):
        print(f"エラー: 指定されたファイル '{doc_path}' が存在しません。")
        return
    
    if not doc_path.lower().endswith('.doc'):
        print(f"エラー: 指定されたファイル '{doc_path}' は.doc形式ではありません。")
        return
    
    # ファイルを変換
    if convert_doc_file(doc_path, encoding=encoding):
        print(f"変換完了: {doc_path}")
    else:
        print(f"変換失敗: {doc_path}")

# メイン処理（直接実行された場合）
if __name__ == "__main__":
    main() 