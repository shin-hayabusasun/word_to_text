#!/usr/bin/env python
# coding: utf-8

import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
from pathlib import Path
import traceback  # トレースバック情報を取得するためのモジュール

# デバッグ用のロギング設定
import logging
logging.basicConfig(level=logging.DEBUG, 
                    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)
logger.info("プログラムを開始します")

# word_to_text_converter.pyからインポート
try:
    from word_to_text_converter import convert_docx_to_text, convert_doc_to_text
    print("モジュールのインポートに成功しました")
    logger.info("モジュールのインポートに成功しました")
except ImportError as e:
    print(f"モジュールのインポートに失敗しました: {e}")
    logger.error(f"モジュールのインポートに失敗しました: {e}")
    traceback.print_exc()
    sys.exit(1)

class WordToTextConverterGUI:
    def __init__(self, root):
        logger.info("GUIクラスの初期化を開始")
        self.root = root
        self.root.title("Word文書テキスト変換ツール")
        self.root.geometry("700x600")
        self.root.resizable(True, True)
        
        # 変数の初期化
        self.directory_path = tk.StringVar()
        self.is_recursive = tk.BooleanVar(value=True)
        self.is_running = False
        self.total_files = 0
        self.processed_files = 0
        self.success_files = 0
        self.failed_files = 0
        self.failed_file_list = []  # 失敗したファイルのリスト
        
        # GUIの設定
        self._setup_ui()
        logger.info("GUIクラスの初期化が完了")
    
    def _setup_ui(self):
        """GUIのレイアウトを設定"""
        # メインフレーム
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # ディレクトリ選択部分
        dir_frame = ttk.LabelFrame(main_frame, text="ディレクトリ選択", padding=5)
        dir_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Entry(dir_frame, textvariable=self.directory_path, width=50).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(dir_frame, text="参照", command=self._browse_directory).pack(side=tk.RIGHT, padx=5)
        
        # オプション部分
        option_frame = ttk.LabelFrame(main_frame, text="オプション", padding=5)
        option_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Checkbutton(option_frame, text="サブディレクトリも含めて変換する", variable=self.is_recursive).pack(anchor=tk.W, padx=5)
        
        # 操作ボタン部分
        button_frame = ttk.Frame(main_frame, padding=5)
        button_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.start_button = ttk.Button(button_frame, text="変換開始", command=self._start_conversion)
        self.start_button.pack(side=tk.LEFT, padx=5)
        
        self.cancel_button = ttk.Button(button_frame, text="キャンセル", command=self._cancel_conversion, state=tk.DISABLED)
        self.cancel_button.pack(side=tk.LEFT, padx=5)
        
        self.save_failed_list_button = ttk.Button(button_frame, text="失敗リスト保存", command=self._save_failed_list, state=tk.DISABLED)
        self.save_failed_list_button.pack(side=tk.LEFT, padx=5)
        
        # 進捗表示部分
        progress_frame = ttk.LabelFrame(main_frame, text="進捗状況", padding=5)
        progress_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, padx=5, pady=5)
        
        self.status_label = ttk.Label(progress_frame, text="準備完了")
        self.status_label.pack(anchor=tk.W, padx=5)
        
        # ログ表示部分
        log_frame = ttk.LabelFrame(main_frame, text="ログ", padding=5)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # ログとエラーのタブ
        tab_control = ttk.Notebook(log_frame)
        tab_control.pack(fill=tk.BOTH, expand=True)
        
        # ログタブ
        log_tab = ttk.Frame(tab_control)
        tab_control.add(log_tab, text="処理ログ")
        
        self.log_text = tk.Text(log_tab, wrap=tk.WORD, height=10)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        log_scrollbar = ttk.Scrollbar(log_tab, command=self.log_text.yview)
        log_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=log_scrollbar.set)
        
        # 失敗リストタブ
        failed_tab = ttk.Frame(tab_control)
        tab_control.add(failed_tab, text="失敗リスト")
        
        self.failed_text = tk.Text(failed_tab, wrap=tk.WORD, height=10)
        self.failed_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        failed_scrollbar = ttk.Scrollbar(failed_tab, command=self.failed_text.yview)
        failed_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.failed_text.config(yscrollcommand=failed_scrollbar.set)
        
        # クレジット表示
        credit_label = ttk.Label(main_frame, text="Created by AI Assistant", foreground="gray")
        credit_label.pack(side=tk.RIGHT, padx=5, pady=5)
        
        logger.info("UI設定が完了しました")
    
    def _browse_directory(self):
        """ディレクトリ選択ダイアログを表示"""
        directory = filedialog.askdirectory(title="マニュアル集のディレクトリを選択")
        if directory:
            self.directory_path.set(directory)
            logger.info(f"ディレクトリが選択されました: {directory}")
    
    def _log(self, message):
        """ログにメッセージを追加"""
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        logger.info(message)
    
    def _update_progress(self, value=None, text=None):
        """進捗状況を更新"""
        if value is not None:
            self.progress_var.set(value)
        
        if text is not None:
            self.status_label.config(text=text)
            logger.info(f"進捗状況を更新: {text}")
    
    def _save_failed_list(self):
        """失敗したファイルのリストをテキストファイルとして保存"""
        if not self.failed_file_list:
            messagebox.showinfo("情報", "失敗したファイルはありません。")
            return
        
        save_path = filedialog.asksaveasfilename(
            title="失敗リストの保存先",
            defaultextension=".txt",
            filetypes=[("テキストファイル", "*.txt"), ("すべてのファイル", "*.*")]
        )
        
        if not save_path:
            return
        
        try:
            with open(save_path, 'w', encoding='utf-8') as f:
                f.write("# 変換に失敗したファイルのリスト\n\n")
                for file in self.failed_file_list:
                    f.write(f"{file}\n")
            
            messagebox.showinfo("保存完了", f"失敗ファイルリストを {save_path} に保存しました。")
            logger.info(f"失敗ファイルリストを {save_path} に保存しました")
        except Exception as e:
            error_msg = f"ファイルの保存中にエラーが発生しました: {str(e)}"
            logger.error(error_msg)
            messagebox.showerror("エラー", error_msg)
    
    def _start_conversion(self):
        """変換処理を開始"""
        directory = self.directory_path.get()
        
        if not directory:
            messagebox.showwarning("警告", "ディレクトリを選択してください。")
            return
        
        if not os.path.exists(directory) or not os.path.isdir(directory):
            error_msg = f"指定されたパス '{directory}' は有効なディレクトリではありません。"
            logger.error(error_msg)
            messagebox.showerror("エラー", error_msg)
            return
        
        # 変換の前に確認
        if not messagebox.askyesno("確認", f"ディレクトリ '{directory}' 内のWord文書をテキストに変換しますか？\n\n" + 
                                 f"再帰的処理: {'有効' if self.is_recursive.get() else '無効'}"):
            return
        
        # UI状態の更新
        self.start_button.config(state=tk.DISABLED)
        self.cancel_button.config(state=tk.NORMAL)
        self.save_failed_list_button.config(state=tk.DISABLED)
        self.is_running = True
        
        # ログと失敗リストの初期化
        self.log_text.delete(1.0, tk.END)
        self.failed_text.delete(1.0, tk.END)
        self._log(f"ディレクトリ '{directory}' 内のWordファイルをテキストに変換します...")
        self._log(f"再帰的処理: {'有効' if self.is_recursive.get() else '無効'}")
        
        # 進捗の初期化
        self.total_files = 0
        self.processed_files = 0
        self.success_files = 0
        self.failed_files = 0
        self.failed_file_list = []
        self._update_progress(0, "ファイル検索中...")
        
        # スレッドで実行
        logger.info(f"変換処理を開始します: {directory}, 再帰的処理: {self.is_recursive.get()}")
        threading.Thread(target=self._convert_files, args=(directory, self.is_recursive.get())).start()
    
    def _cancel_conversion(self):
        """変換処理をキャンセル"""
        if self.is_running:
            if messagebox.askyesno("確認", "変換処理をキャンセルしますか？"):
                self.is_running = False
                self._log("変換処理をキャンセルしました。")
                self._update_progress(text="キャンセルされました")
                self._reset_ui()
                logger.info("変換処理がキャンセルされました")
    
    def _reset_ui(self):
        """UI状態をリセット"""
        self.start_button.config(state=tk.NORMAL)
        self.cancel_button.config(state=tk.DISABLED)
        if self.failed_file_list:
            self.save_failed_list_button.config(state=tk.NORMAL)
        logger.info("UI状態をリセットしました")
    
    def _add_to_failed_list(self, file_path, error_message=None):
        """失敗したファイルをリストに追加"""
        self.failed_file_list.append(file_path)
        
        # 失敗リストのテキストに追加
        text = f"{file_path}"
        if error_message:
            text += f"\n  -> 原因: {error_message}"
        text += "\n\n"
        
        self.root.after(0, lambda t=text: self.failed_text.insert(tk.END, t))
        self.root.after(0, lambda: self.failed_text.see(tk.END))
        logger.warning(f"失敗リストに追加: {file_path}, 原因: {error_message}")
    
    def _convert_files(self, directory, recursive):
        """ディレクトリ内のファイルを変換（スレッドで実行）"""
        try:
            # 絶対パスに変換
            directory = os.path.abspath(directory)
            
            # 再帰的検索パターン
            pattern = '**/*' if recursive else '*'
            
            # ファイル数をカウント
            docx_files = list(Path(directory).glob(f"{pattern}.docx"))
            doc_files = list(Path(directory).glob(f"{pattern}.doc"))
            
            self.total_files = len(docx_files) + len(doc_files)
            
            if self.total_files == 0:
                self.root.after(0, lambda: self._log("変換対象のファイルが見つかりませんでした。"))
                self.root.after(0, lambda: self._update_progress(100, "完了（対象ファイルなし）"))
                self.root.after(0, self._reset_ui)
                logger.info("変換対象のファイルが見つかりませんでした")
                return
            
            self.root.after(0, lambda: self._update_progress(text=f"合計 {self.total_files} ファイルを処理します..."))
            logger.info(f"合計 {self.total_files} ファイルを処理します")
            
            # .docxファイルを処理
            for docx_file in docx_files:
                if not self.is_running:
                    break
                
                docx_file_str = str(docx_file)
                self.root.after(0, lambda f=docx_file_str: self._log(f"処理中: {f}"))
                logger.info(f"処理中: {docx_file_str}")
                
                try:
                    output_path = convert_docx_to_text(docx_file_str)
                    if output_path:
                        self.success_files += 1
                        self.root.after(0, lambda p=output_path: self._log(f"  変換完了: {p}"))
                        logger.info(f"変換完了: {output_path}")
                    else:
                        self.failed_files += 1
                        self._add_to_failed_list(docx_file_str, "変換失敗 (詳細はログを確認)")
                        logger.error(f"変換失敗: {docx_file_str}")
                except Exception as e:
                    error_msg = str(e)
                    self.failed_files += 1
                    self.root.after(0, lambda f=docx_file_str, err=error_msg: self._log(f"  エラー（{f}）: {err}"))
                    self._add_to_failed_list(docx_file_str, error_msg)
                    logger.error(f"エラー（{docx_file_str}）: {error_msg}", exc_info=True)
                
                self.processed_files += 1
                progress = (self.processed_files / self.total_files) * 100
                self.root.after(0, lambda p=progress: self._update_progress(p, f"処理中... ({self.processed_files}/{self.total_files})"))
            
            # .docファイルを処理
            for doc_file in doc_files:
                if not self.is_running:
                    break
                
                doc_file_str = str(doc_file)
                self.root.after(0, lambda f=doc_file_str: self._log(f"処理中: {f}"))
                logger.info(f"処理中: {doc_file_str}")
                
                try:
                    output_path = convert_doc_to_text(doc_file_str)
                    if output_path:
                        self.success_files += 1
                        self.root.after(0, lambda p=output_path: self._log(f"  変換完了: {p}"))
                        logger.info(f"変換完了: {output_path}")
                    else:
                        self.failed_files += 1
                        self._add_to_failed_list(doc_file_str, "変換失敗 (詳細はログを確認)")
                        logger.error(f"変換失敗: {doc_file_str}")
                except Exception as e:
                    error_msg = str(e)
                    self.failed_files += 1
                    self.root.after(0, lambda f=doc_file_str, err=error_msg: self._log(f"  エラー（{f}）: {err}"))
                    self._add_to_failed_list(doc_file_str, error_msg)
                    logger.error(f"エラー（{doc_file_str}）: {error_msg}", exc_info=True)
                
                self.processed_files += 1
                progress = (self.processed_files / self.total_files) * 100
                self.root.after(0, lambda p=progress: self._update_progress(p, f"処理中... ({self.processed_files}/{self.total_files})"))
            
            # 完了メッセージ
            if self.is_running:
                summary = f"変換処理が完了しました。成功: {self.success_files}, 失敗: {self.failed_files}, 合計: {self.processed_files}/{self.total_files}"
                self.root.after(0, lambda: self._log(summary))
                self.root.after(0, lambda: self._update_progress(100, "完了"))
                logger.info(summary)
                
                # 失敗したファイルがある場合の表示
                if self.failed_files > 0:
                    self.root.after(0, lambda: self._log(f"\n失敗したファイルが {self.failed_files} 件あります。「失敗リスト」タブで確認できます。"))
                    self.root.after(0, lambda: messagebox.showinfo("完了", f"{summary}\n\n失敗したファイルが {self.failed_files} 件あります。「失敗リスト」タブで確認してください。"))
                    logger.warning(f"失敗したファイルが {self.failed_files} 件あります")
                else:
                    self.root.after(0, lambda: messagebox.showinfo("完了", summary))
            
        except Exception as e:
            error_msg = f"エラーが発生しました: {str(e)}"
            self.root.after(0, lambda err=error_msg: self._log(error_msg))
            self.root.after(0, lambda: self._update_progress(text="エラーが発生しました"))
            logger.error(error_msg, exc_info=True)
        
        finally:
            self.is_running = False
            self.root.after(0, self._reset_ui)

    def convert_single_file(self, file_path, output_dir=None):
        """単一のファイルを変換する"""
        try:
            file_path = Path(file_path)
            file_ext = file_path.suffix.lower()
            
            if output_dir:
                output_dir = Path(output_dir)
                # 出力ディレクトリが存在しない場合は作成
                if not output_dir.exists():
                    output_dir.mkdir(parents=True)
                output_path = output_dir / f"{file_path.stem}.txt"
            else:
                output_path = file_path.with_suffix('.txt')
            
            logger.info(f"変換開始: {file_path} -> {output_path}")
            self._update_progress(text=f"変換中: {file_path.name}")
            
            if file_ext == '.docx':
                result_path = convert_docx_to_text(str(file_path), str(output_path))
                logger.info(f"DOCX変換完了: {result_path}")
                self._update_progress(text=f"変換完了: {file_path.name}")
                return result_path
            elif file_ext == '.doc':
                result_path = convert_doc_to_text(str(file_path), str(output_path))
                logger.info(f"DOC変換完了: {result_path}")
                
                # 変換結果の検証
                if os.path.exists(output_path):
                    with open(output_path, 'r', encoding='utf-8') as f:
                        content = f.read()
                    if len(content) < 100:  # 内容が少なすぎる場合は警告
                        logger.warning(f"変換結果のテキストが少なすぎます: {len(content)}文字")
                        self._update_progress(text=f"警告: {file_path.name} の変換結果が短すぎます")
                        # 強化版日本語抽出を再試行
                        try:
                            logger.info("強化版日本語テキスト抽出を再試行します")
                            result_path = extract_japanese_text_enhanced(str(file_path), str(output_path))
                            logger.info(f"再変換完了: {result_path}")
                        except Exception as retry_err:
                            logger.error(f"再変換失敗: {str(retry_err)}")
                    else:
                        logger.info(f"変換結果のテキスト: {len(content)}文字")
                        self._update_progress(text=f"変換完了: {file_path.name}")
                else:
                    logger.error(f"出力ファイルが作成されませんでした: {output_path}")
                    self._update_progress(text=f"エラー: {file_path.name} の出力ファイルが作成されませんでした")
                
                return result_path
            else:
                logger.error(f"サポートされていないファイル形式: {file_ext}")
                self._update_progress(text=f"エラー: サポートされていないファイル形式 {file_ext}")
                messagebox.showerror("エラー", f"サポートされていないファイル形式: {file_ext}")
                return None
        except Exception as e:
            logger.error(f"変換エラー: {str(e)}")
            logger.error(traceback.format_exc())
            self._update_progress(text=f"エラー: {file_path.name} の変換に失敗しました")
            messagebox.showerror("変換エラー", f"ファイル {file_path.name} の変換中にエラーが発生しました:\n{str(e)}")
            return None

def main():
    try:
        logger.info("GUIを起動します...")
        print("GUIを起動します...")
        root = tk.Tk()
        app = WordToTextConverterGUI(root)
        logger.info("mainloopを開始します...")
        print("mainloopを開始します...")
        root.mainloop()
    except Exception as e:
        error_msg = f"エラーが発生しました: {str(e)}"
        print(error_msg)
        logger.error(error_msg, exc_info=True)
        traceback.print_exc()
        return 1
    return 0

if __name__ == "__main__":
    try:
        logger.info("プログラムを開始します...")
        print("プログラムを開始します...")
        sys.exit(main())
    except Exception as e:
        error_msg = f"予期せぬエラーが発生しました: {str(e)}"
        print(error_msg)
        logger.error(error_msg, exc_info=True)
        traceback.print_exc()
        sys.exit(1) 