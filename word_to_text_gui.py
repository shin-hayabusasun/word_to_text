#!/usr/bin/env python
# coding: utf-8

import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
from pathlib import Path
import traceback  # トレースバック情報を取得するためのモジュール
import tkinterdnd2  # ドラッグ&ドロップのためのライブラリ

# デバッグ用のロギング設定
import logging
logging.basicConfig(level=logging.DEBUG, 
                    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)
logger.info("プログラムを開始します")

# word_to_text_converter.pyからインポート
try:
    from word_to_text_converter import convert_docx_to_text, convert_doc_to_text, process_directory, extract_japanese_text_enhanced
    print("モジュールのインポートに成功しました")
    logger.info("モジュールのインポートに成功しました")
except ImportError as e:
    print(f"モジュールのインポートに失敗しました: {e}")
    logger.error(f"モジュールのインポートに失敗しました: {e}")
    # tkinterDnD2がない場合は標準のtkinterを使用
    try:
        import tkinter as tkinterdnd2
        logger.warning("tkinterdnd2をインポートできませんでした。標準のtkinterを使用します")
    except:
        pass
    traceback.print_exc()
    # 致命的なエラーではないので、継続
    pass

class WordToTextConverterGUI:
    def __init__(self, root):
        logger.info("GUIクラスの初期化を開始")
        self.root = root
        self.root.title("Word文書テキスト変換ツール")
        self.root.geometry("700x600")
        self.root.resizable(True, True)
        
        # アイコンの設定（あれば）
        try:
            self.root.iconbitmap("icon.ico")
        except:
            pass
        
        # 変数の初期化
        self.directory_path = tk.StringVar()
        self.is_recursive = tk.BooleanVar(value=True)
        self.force_utf8 = tk.BooleanVar(value=True)  # UTF-8優先フラグ
        self.use_sjis = tk.BooleanVar(value=False)   # Shift-JIS優先フラグ
        self.is_running = False
        self.total_files = 0
        self.processed_files = 0
        self.success_files = 0
        self.failed_files = 0
        self.failed_file_list = []  # 失敗したファイルのリスト
        
        # GUIの設定
        self._setup_ui()
        
        # ドラッグ&ドロップの設定
        try:
            self.root.drop_target_register(tkinterdnd2.DND_FILES)
            self.root.dnd_bind('<<Drop>>', self._on_drop)
            logger.info("ドラッグ&ドロップの設定が完了")
        except Exception as e:
            logger.error(f"ドラッグ&ドロップの設定に失敗しました: {e}")
            # ドラッグ&ドロップは必須ではないので、エラーでも続行
        
        logger.info("GUIクラスの初期化が完了")
    
    def _on_drop(self, event):
        """ファイルがドロップされたときの処理"""
        try:
            # ドロップされたファイルパス（複数ファイルの場合はスペース区切り）
            data = event.data
            # Windows環境では {} で囲まれているので除去
            if data.startswith('{') and data.endswith('}'):
                data = data[1:-1]
            
            # 複数ファイル対応（スペース区切り）
            files = data.split()
            
            # 単一ファイルの場合はそのまま処理
            if len(files) == 1:
                path = files[0]
                if os.path.isdir(path):
                    self.directory_path.set(path)
                    logger.info(f"ディレクトリがドロップされました: {path}")
                elif path.lower().endswith(('.doc', '.docx')):
                    self._process_single_file(path)
                    logger.info(f"ファイルがドロップされました: {path}")
                else:
                    messagebox.showwarning("警告", "サポートされていないファイル形式です。\n.docまたは.docxファイルを選択してください。")
            else:
                # 複数ファイルの処理
                self._process_multiple_files(files)
        except Exception as e:
            error_msg = f"ドロップ処理中にエラーが発生しました: {str(e)}"
            logger.error(error_msg)
            messagebox.showerror("エラー", error_msg)
    
    def _process_single_file(self, file_path):
        """単一ファイルを処理"""
        if not os.path.exists(file_path):
            messagebox.showerror("エラー", f"ファイル '{file_path}' が存在しません。")
            return
        
        # 処理前の確認
        if not messagebox.askyesno("確認", f"ファイル '{os.path.basename(file_path)}' をテキストに変換しますか？\n\n" + 
                                  f"UTF-8優先: {'有効' if self.force_utf8.get() else '無効'}\n" + 
                                  f"Shift-JIS優先: {'有効' if self.use_sjis.get() else '無効'}"):
            return
        
        # UI状態の更新
        self.start_button.config(state=tk.DISABLED)
        self.cancel_button.config(state=tk.NORMAL)
        self.save_failed_list_button.config(state=tk.DISABLED)
        self.is_running = True
        
        # ログの初期化
        self.log_text.delete(1.0, tk.END)
        self.failed_text.delete(1.0, tk.END)
        self._log(f"ファイル '{file_path}' をテキストに変換します...")
        self._log(f"UTF-8優先: {'有効' if self.force_utf8.get() else '無効'}")
        self._log(f"Shift-JIS優先: {'有効' if self.use_sjis.get() else '無効'}")
        
        # 進捗の初期化
        self.total_files = 1
        self.processed_files = 0
        self.success_files = 0
        self.failed_files = 0
        self.failed_file_list = []
        self._update_progress(0, "処理中...")
        
        # スレッドで実行
        threading.Thread(target=self._convert_single_file, args=(file_path,)).start()
    
    def _process_multiple_files(self, files):
        """複数のファイルを処理"""
        # Word文書だけをフィルタリング
        word_files = [f for f in files if f.lower().endswith(('.doc', '.docx'))]
        
        if not word_files:
            messagebox.showwarning("警告", "変換可能なWord文書が見つかりませんでした。")
            return
        
        # 処理前の確認
        if not messagebox.askyesno("確認", f"{len(word_files)} 個のWord文書をテキストに変換しますか？\n\n" + 
                                 f"UTF-8優先: {'有効' if self.force_utf8.get() else '無効'}\n" + 
                                 f"Shift-JIS優先: {'有効' if self.use_sjis.get() else '無効'}"):
            return
        
        # UI状態の更新
        self.start_button.config(state=tk.DISABLED)
        self.cancel_button.config(state=tk.NORMAL)
        self.save_failed_list_button.config(state=tk.DISABLED)
        self.is_running = True
        
        # ログの初期化
        self.log_text.delete(1.0, tk.END)
        self.failed_text.delete(1.0, tk.END)
        self._log(f"{len(word_files)} 個のファイルをテキストに変換します...")
        self._log(f"UTF-8優先: {'有効' if self.force_utf8.get() else '無効'}")
        self._log(f"Shift-JIS優先: {'有効' if self.use_sjis.get() else '無効'}")
        
        # 進捗の初期化
        self.total_files = len(word_files)
        self.processed_files = 0
        self.success_files = 0
        self.failed_files = 0
        self.failed_file_list = []
        self._update_progress(0, "処理中...")
        
        # スレッドで実行
        threading.Thread(target=self._convert_multiple_files, args=(word_files,)).start()
    
    def _convert_single_file(self, file_path):
        """単一ファイルの変換をスレッドで実行"""
        try:
            if file_path.lower().endswith('.docx'):
                output_path = convert_docx_to_text(file_path)
                if output_path:
                    self.success_files += 1
                    self.root.after(0, lambda: self._log(f"変換完了: {output_path}"))
                else:
                    self.failed_files += 1
                    self._add_to_failed_list(file_path, "変換失敗")
            elif file_path.lower().endswith('.doc'):
                output_path = convert_doc_to_text(file_path, force_utf8=self.force_utf8.get(), use_sjis=self.use_sjis.get())
                if output_path:
                    self.success_files += 1
                    self.root.after(0, lambda: self._log(f"変換完了: {output_path}"))
                else:
                    self.failed_files += 1
                    self._add_to_failed_list(file_path, "変換失敗")
        except Exception as e:
            error_msg = str(e)
            self.failed_files += 1
            self.root.after(0, lambda: self._log(f"エラー（{file_path}）: {error_msg}"))
            self._add_to_failed_list(file_path, error_msg)
        
        self.processed_files += 1
        progress = (self.processed_files / self.total_files) * 100
        self.root.after(0, lambda: self._update_progress(progress, f"完了"))
        
        # 完了メッセージ
        if self.is_running:
            summary = f"変換処理が完了しました。成功: {self.success_files}, 失敗: {self.failed_files}"
            self.root.after(0, lambda: self._log(summary))
            if self.failed_files > 0:
                self.root.after(0, lambda: messagebox.showinfo("完了", f"{summary}\n\n失敗したファイルがあります。詳細はログを確認してください。"))
            else:
                self.root.after(0, lambda: messagebox.showinfo("完了", summary))
        
        self.is_running = False
        self.root.after(0, self._reset_ui)
    
    def _convert_multiple_files(self, files):
        """複数ファイルの変換をスレッドで実行"""
        for file_path in files:
            if not self.is_running:
                break
            
            self.root.after(0, lambda f=file_path: self._log(f"処理中: {f}"))
            
            try:
                if file_path.lower().endswith('.docx'):
                    output_path = convert_docx_to_text(file_path)
                    if output_path:
                        self.success_files += 1
                        self.root.after(0, lambda p=output_path: self._log(f"  変換完了: {p}"))
                    else:
                        self.failed_files += 1
                        self._add_to_failed_list(file_path, "変換失敗")
                elif file_path.lower().endswith('.doc'):
                    output_path = convert_doc_to_text(file_path, force_utf8=self.force_utf8.get(), use_sjis=self.use_sjis.get())
                    if output_path:
                        self.success_files += 1
                        self.root.after(0, lambda p=output_path: self._log(f"  変換完了: {p}"))
                    else:
                        self.failed_files += 1
                        self._add_to_failed_list(file_path, "変換失敗")
            except Exception as e:
                error_msg = str(e)
                self.failed_files += 1
                self.root.after(0, lambda f=file_path, err=error_msg: self._log(f"  エラー（{f}）: {err}"))
                self._add_to_failed_list(file_path, error_msg)
            
            self.processed_files += 1
            progress = (self.processed_files / self.total_files) * 100
            self.root.after(0, lambda p=progress: self._update_progress(p, f"処理中... ({self.processed_files}/{self.total_files})"))
        
        # 完了メッセージ
        if self.is_running:
            summary = f"変換処理が完了しました。成功: {self.success_files}, 失敗: {self.failed_files}, 合計: {self.total_files}"
            self.root.after(0, lambda: self._log(summary))
            self.root.after(0, lambda: self._update_progress(100, "完了"))
            
            if self.failed_files > 0:
                self.root.after(0, lambda: self._log(f"\n失敗したファイルが {self.failed_files} 件あります。「失敗リスト」タブで確認できます。"))
                self.root.after(0, lambda: messagebox.showinfo("完了", f"{summary}\n\n失敗したファイルが {self.failed_files} 件あります。「失敗リスト」タブで確認してください。"))
            else:
                self.root.after(0, lambda: messagebox.showinfo("完了", summary))
        
        self.is_running = False
        self.root.after(0, self._reset_ui)
    
    def _setup_ui(self):
        """GUIのレイアウトを設定"""
        # メインフレーム
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # ファイル/ディレクトリ選択部分
        file_frame = ttk.LabelFrame(main_frame, text="ファイル/ディレクトリ選択", padding=5)
        file_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Entry(file_frame, textvariable=self.directory_path, width=50).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        button_frame = ttk.Frame(file_frame)
        button_frame.pack(side=tk.RIGHT, padx=5)
        
        ttk.Button(button_frame, text="ファイル選択", command=self._browse_file).pack(side=tk.LEFT, padx=2)
        ttk.Button(button_frame, text="フォルダ選択", command=self._browse_directory).pack(side=tk.LEFT, padx=2)
        
        # オプション部分
        option_frame = ttk.LabelFrame(main_frame, text="オプション", padding=5)
        option_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Checkbutton(option_frame, text="サブディレクトリも含めて変換する", variable=self.is_recursive).pack(anchor=tk.W, padx=5, pady=2)
        ttk.Checkbutton(option_frame, text="UTF-8エンコーディングを優先する（日本語特化）", variable=self.force_utf8).pack(anchor=tk.W, padx=5, pady=2)
        ttk.Checkbutton(option_frame, text="Shift-JISエンコーディングを優先する", variable=self.use_sjis).pack(anchor=tk.W, padx=5, pady=2)
        
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
        
        self.status_label = ttk.Label(progress_frame, text="準備完了 - ファイルをここにドロップすることもできます")
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
        
        # ガイド表示
        guide_frame = ttk.Frame(main_frame)
        guide_frame.pack(fill=tk.X, padx=5, pady=5)
        
        guide_text = "ドラッグ&ドロップ: Word文書またはフォルダをここにドロップできます"
        ttk.Label(guide_frame, text=guide_text, foreground="gray").pack(side=tk.LEFT, padx=5)
        
        # クレジット表示
        credit_label = ttk.Label(main_frame, text="Created by AI Assistant", foreground="gray")
        credit_label.pack(side=tk.RIGHT, padx=5, pady=5)
        
        logger.info("UI設定が完了しました")
    
    def _browse_file(self):
        """ファイル選択ダイアログを表示"""
        file_path = filedialog.askopenfilename(
            title="Word文書を選択",
            filetypes=[("Word文書", "*.doc;*.docx"), ("すべてのファイル", "*.*")]
        )
        if file_path:
            self.directory_path.set(file_path)
            logger.info(f"ファイルが選択されました: {file_path}")
    
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
        path = self.directory_path.get()
        
        if not path:
            messagebox.showwarning("警告", "ファイルまたはディレクトリを選択してください。")
            return
        
        if not os.path.exists(path):
            error_msg = f"指定されたパス '{path}' は存在しません。"
            logger.error(error_msg)
            messagebox.showerror("エラー", error_msg)
            return
        
        # ファイルかディレクトリかを判断
        if os.path.isdir(path):
            # ディレクトリの処理
            if not messagebox.askyesno("確認", 
                                    f"ディレクトリ '{path}' 内のWord文書をテキストに変換しますか？\n\n" + 
                                    f"再帰的処理: {'有効' if self.is_recursive.get() else '無効'}\n" + 
                                    f"UTF-8優先: {'有効' if self.force_utf8.get() else '無効'}\n" + 
                                    f"Shift-JIS優先: {'有効' if self.use_sjis.get() else '無効'}"):
                return
            
            # UI状態の更新
            self.start_button.config(state=tk.DISABLED)
            self.cancel_button.config(state=tk.NORMAL)
            self.save_failed_list_button.config(state=tk.DISABLED)
            self.is_running = True
            
            # ログと失敗リストの初期化
            self.log_text.delete(1.0, tk.END)
            self.failed_text.delete(1.0, tk.END)
            self._log(f"ディレクトリ '{path}' 内のWordファイルをテキストに変換します...")
            self._log(f"再帰的処理: {'有効' if self.is_recursive.get() else '無効'}")
            self._log(f"UTF-8優先: {'有効' if self.force_utf8.get() else '無効'}")
            self._log(f"Shift-JIS優先: {'有効' if self.use_sjis.get() else '無効'}")
            
            # 進捗の初期化
            self.total_files = 0
            self.processed_files = 0
            self.success_files = 0
            self.failed_files = 0
            self.failed_file_list = []
            self._update_progress(0, "ファイル検索中...")
            
            # スレッドで実行
            logger.info(f"変換処理を開始します: {path}, 再帰的処理: {self.is_recursive.get()}, UTF-8優先: {self.force_utf8.get()}, Shift-JIS優先: {self.use_sjis.get()}")
            threading.Thread(target=self._convert_directory, args=(path, self.is_recursive.get(), self.force_utf8.get(), self.use_sjis.get())).start()
        else:
            # 単一ファイルの処理
            self._process_single_file(path)
    
    def _convert_directory(self, directory_path, recursive, force_utf8, use_sjis=False):
        """ディレクトリ内のファイルを変換（スレッドで実行）"""
        try:
            # process_directory関数を使用して変換
            success_files, failed_files = process_directory(directory_path, recursive, force_utf8, use_sjis)
            
            # 結果を更新
            self.success_files = len(success_files)
            self.failed_files = len(failed_files)
            self.total_files = self.success_files + self.failed_files
            self.processed_files = self.total_files
            self.failed_file_list = failed_files
            
            # 失敗したファイルを失敗リストに表示
            for file in failed_files:
                self._add_to_failed_list(file)
            
            # 完了メッセージ
            summary = f"変換処理が完了しました。成功: {self.success_files}, 失敗: {self.failed_files}, 合計: {self.total_files}"
            self.root.after(0, lambda: self._log(summary))
            self.root.after(0, lambda: self._update_progress(100, "完了"))
            
            if self.failed_files > 0:
                self.root.after(0, lambda: self._log(f"\n失敗したファイルが {self.failed_files} 件あります。「失敗リスト」タブで確認できます。"))
                self.root.after(0, lambda: messagebox.showinfo("完了", f"{summary}\n\n失敗したファイルが {self.failed_files} 件あります。「失敗リスト」タブで確認してください。"))
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

def main():
    try:
        logger.info("GUIを起動します...")
        print("GUIを起動します...")
        
        # tkinterdnd2がインポートできた場合はそちらを使用
        try:
            root = tkinterdnd2.Tk()
        except:
            # 通常のtkinterにフォールバック
            root = tk.Tk()
            logger.warning("tkinterdnd2が使用できないため、標準のtkinterを使用します")
        
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