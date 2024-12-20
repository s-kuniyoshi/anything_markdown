import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
from markitdown import MarkItDown

try:
    from openai import OpenAI
    OPENAI_AVAILABLE = True
except ImportError:
    OPENAI_AVAILABLE = False

from markitdown._markitdown import UnsupportedFormatException

import pandas as pd
import datetime
import tkinter.ttk as ttk
import tkinter.font as tkFont
from PIL import Image, ImageTk, __version__ as PILLOW_VERSION  # Pillowライブラリを使用してアイコンを表示

class MarkdownConverterGUI:
    SUPPORTED_EXTENSIONS = [
        '.pdf', '.ppt', '.doc', '.docx', '.xls', '.xlsx',
        '.jpg', '.jpeg', '.png', '.html', '.csv', '.json',
        '.xml', '.zip', '.txt', '.pptx'
    ]

    def __init__(self, root):
        self.root = root
        self.root.title("MarkItDown GUI Converter")
        self.root.geometry("800x750")  # 初期サイズを設定
        self.root.resizable(True, True)  # ウィンドウのリサイズを許可
        self.use_chatgpt = tk.BooleanVar()
        self.show_api_key = False  # APIキー表示状態を管理

        # フォント設定
        self.style = ttk.Style()
        self.style.theme_use('clam')

        # フォントのリスト
        preferred_fonts = ["IPAPGothic", "Noto Sans CJK JP", "TakaoGothic", "VL Gothic"]
        available_fonts = list(tkFont.families())
        selected_font = None
        for font in preferred_fonts:
            if font in available_fonts:
                selected_font = font
                break
        if not selected_font:
            selected_font = "Helvetica"  # デフォルトフォント

        # スタイルの設定
        self.style.configure("TButton",
                             padding=6,
                             relief="flat",
                             background="#4CAF50",
                             foreground="white",
                             font=(selected_font, 10, "bold"))
        self.style.map("TButton",
                       background=[('active', '#45a049')])

        self.style.configure("TCheckbutton",
                             font=(selected_font, 10))
        self.style.configure("TLabel",
                             font=(selected_font, 10))
        self.style.configure("TEntry",
                             font=(selected_font, 10))

        # Input Directory Selection
        self.input_label = ttk.Label(root, text="入力ディレクトリ:")
        self.input_label.pack(pady=(20, 5))
        self.input_entry = ttk.Entry(root, width=100)
        self.input_entry.pack(pady=5)
        self.input_button = ttk.Button(root, text="選択", command=self.select_input_directory)
        self.input_button.pack(pady=5)

        # Output Directory Selection
        self.output_label = ttk.Label(root, text="出力ディレクトリ:")
        self.output_label.pack(pady=(20, 5))
        self.output_entry = ttk.Entry(root, width=100)
        self.output_entry.pack(pady=5)
        self.output_button = ttk.Button(root, text="選択", command=self.select_output_directory)
        self.output_button.pack(pady=5)

        # Pillowのバージョンに応じてリサンプリングフィルターを選択
        pillow_version = tuple(map(int, PILLOW_VERSION.split('.')[:2]))
        if pillow_version >= (10, 0):
            resample_filter = Image.Resampling.LANCZOS
        else:
            resample_filter = Image.ANTIALIAS

        # ChatGPTオプション用のフレームを作成
        self.chatgpt_frame = ttk.Frame(root)
        self.chatgpt_frame.pack(pady=20, fill='x', padx=20)

        # Use ChatGPT Checkbox
        if OPENAI_AVAILABLE:
            self.chatgpt_checkbox = ttk.Checkbutton(
                self.chatgpt_frame, 
                text="画像の説明にChatGPTを使用する", 
                variable=self.use_chatgpt, 
                command=self.toggle_api_key_entry
            )
        else:
            self.chatgpt_checkbox = ttk.Checkbutton(
                self.chatgpt_frame,
                text="画像の説明にChatGPTを使用する (openai未インストール)",
                state='disabled'
            )
        self.chatgpt_checkbox.pack(anchor='w')

        # API Key Entry (Initially hidden)
        self.api_key_frame = ttk.Frame(self.chatgpt_frame)
        
        # 一時的に背景色を設定してフレームの存在を確認
        self.style.configure('APIFRAME.TFrame', background='lightgrey')
        self.api_key_frame.configure(style='APIFRAME.TFrame')

        # APIキーラベル
        self.api_key_label = ttk.Label(self.api_key_frame, text="OpenAI APIキー:")
        self.api_key_label.pack(side=tk.LEFT, padx=(0, 10), pady=5)

        # APIキーエントリー
        self.api_key_entry = ttk.Entry(self.api_key_frame, width=50, show="*")  # 入力をマスク
        self.api_key_entry.pack(side=tk.LEFT, pady=5)

        # アイコンボタンの作成（表示/非表示切替）
        # アイコン画像のパスを指定（imagesフォルダ内にeye_closed.pngとeye_open.pngを配置）
        try:
            with Image.open("images/eye_closed.png") as img:
                closed_eye_image = img.resize((20, 20), resample_filter)
                self.icon_closed = ImageTk.PhotoImage(closed_eye_image)
        except Exception as e:
            messagebox.showerror("エラー", f"eye_closed.png を読み込めませんでした: {e}")
            self.icon_closed = None

        try:
            with Image.open("images/eye_open.png") as img:
                open_eye_image = img.resize((20, 20), resample_filter)
                self.icon_open = ImageTk.PhotoImage(open_eye_image)
        except Exception as e:
            messagebox.showerror("エラー", f"eye_open.png を読み込めませんでした: {e}")
            self.icon_open = None

        self.toggle_button = ttk.Button(
            self.api_key_frame, 
            image=self.icon_closed if self.icon_closed else '', 
            command=self.toggle_api_key_visibility
        )
        self.toggle_button.pack(side=tk.LEFT, padx=(5, 0), pady=5)

        # 初期状態では非表示
        self.api_key_frame.pack_forget()

        # Convert Button
        self.convert_button = ttk.Button(root, text="変換開始", command=self.start_conversion)
        self.convert_button.pack(pady=10)

        # Progress Bar
        self.progress = ttk.Progressbar(root, orient='horizontal', length=700, mode='determinate')
        self.progress.pack(pady=10)

        # Status Text
        self.status_text = tk.Text(
            root,
            height=15,
            width=100,
            state='disabled',
            bg="#f0f0f0",
            fg="#333333",
            font=(selected_font, 10)
        )
        self.status_text.pack(pady=10)

        # Initialize log file
        self.initialize_log_file()

    def toggle_api_key_entry(self):
        if self.use_chatgpt.get():
            self.api_key_frame.pack(pady=5, fill='x')
            self.log_status("ChatGPT 使用が有効になりました。APIキー入力フィールドを表示します。")
        else:
            self.api_key_frame.pack_forget()
            self.log_status("ChatGPT 使用が無効になりました。APIキー入力フィールドを非表示にします。")

    def toggle_api_key_visibility(self):
        if self.show_api_key:
            self.api_key_entry.config(show="*")
            if self.icon_closed:
                self.toggle_button.config(image=self.icon_closed)
            self.show_api_key = False
            self.log_status("APIキーの表示を隠しました。")
        else:
            self.api_key_entry.config(show="")
            if self.icon_open:
                self.toggle_button.config(image=self.icon_open)
            self.show_api_key = True
            self.log_status("APIキーを表示しました。")

    def initialize_log_file(self):
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        logs_dir = os.path.join(os.getcwd(), "logs")
        os.makedirs(logs_dir, exist_ok=True)
        self.log_file_path = os.path.join(logs_dir, f"conversion_log_{timestamp}.txt")
        with open(self.log_file_path, 'w', encoding='utf-8') as log_file:
            log_file.write("Markdown変換ログ\n")
            log_file.write("================\n")

    def select_input_directory(self):
        directory = filedialog.askdirectory(initialdir="/host")
        if directory:
            self.input_entry.delete(0, tk.END)
            self.input_entry.insert(0, directory)
            self.log_status(f"入力ディレクトリを選択しました: {directory}")

    def select_output_directory(self):
        directory = filedialog.askdirectory(initialdir="/host")
        if directory:
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, directory)
            self.log_status(f"出力ディレクトリを選択しました: {directory}")

    def log_status(self, message):
        self.status_text.config(state='normal')
        self.status_text.insert(tk.END, message + "\n")
        self.status_text.see(tk.END)
        self.status_text.config(state='disabled')
        self.root.update_idletasks()
        # ログファイルにも書き込む
        with open(self.log_file_path, 'a', encoding='utf-8') as log_file:
            log_file.write(message + "\n")
        print(message)  # コンソールにも出力

    def start_conversion(self):
        input_dir = self.input_entry.get()
        output_dir = self.output_entry.get()
        use_chatgpt = self.use_chatgpt.get()
        api_key = self.api_key_entry.get() if use_chatgpt else None  # ChatGPT使用時のみAPIキーを取得

        # バリデーション
        if not input_dir or not os.path.isdir(input_dir):
            messagebox.showerror("エラー", "有効な入力ディレクトリを選択してください。")
            self.log_status("エラー: 有効な入力ディレクトリが選択されていません。")
            return
        if not output_dir:
            messagebox.showerror("エラー", "出力ディレクトリを指定してください。")
            self.log_status("エラー: 出力ディレクトリが指定されていません。")
            return
        if use_chatgpt and not OPENAI_AVAILABLE:
            messagebox.showerror("エラー", "ChatGPTを使用するには `openai` ライブラリをインストールしてください。")
            self.log_status("エラー: ChatGPTを使用するための `openai` ライブラリがインストールされていません。")
            return
        if use_chatgpt and not api_key:
            messagebox.showerror("エラー", "ChatGPTを使用するにはOpenAIのAPIキーを入力してください。")
            self.log_status("エラー: ChatGPT使用時にAPIキーが入力されていません。")
            return

        # LLMのセットアップ
        if use_chatgpt:
            try:
                client = OpenAI(api_key=api_key)  # APIキーを渡す
                md_converter = MarkItDown(llm_client=client, llm_model="gpt-4")
                self.log_status("ChatGPT クライアントを初期化しました。")
            except Exception as e:
                messagebox.showerror("エラー", f"ChatGPTの初期化に失敗しました: {e}")
                self.log_status(f"エラー: ChatGPTの初期化に失敗しました: {e}")
                return
        else:
            md_converter = MarkItDown()
            self.log_status("ChatGPTを使用せずに変換を開始します。")

        self.log_status("変換を開始します...")
        self.root.update()

        # ディレクトリを処理
        try:
            self.process_directory(md_converter, input_dir, output_dir)
            self.log_status("全てのファイルの変換が完了しました。")
            messagebox.showinfo("完了", "全てのファイルの変換が完了しました。")
        except Exception as e:
            self.log_status(f"変換中に重大なエラーが発生しました: {e}")
            messagebox.showerror("エラー", f"変換中に重大なエラーが発生しました: {e}")

    def convert_file(self, md_converter, input_path, output_path):
        try:
            # .xls ファイルの場合、.xlsx に変換
            base, ext = os.path.splitext(input_path)
            if ext.lower() == '.xls':
                self.log_status(f"🔄 .xls ファイルを .xlsx に変換中: {input_path}")
                input_path = self.convert_xls_to_xlsx(input_path)
                if not input_path:
                    self.log_status(f"❌ .xls ファイルの変換に失敗: {input_path}")
                    return

            result = md_converter.convert(input_path)
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(result.text_content)
            self.log_status(f"✅ 変換成功: {input_path} -> {output_path}")
        except UnsupportedFormatException:
            self.log_status(f"❌ サポートされていない形式: {input_path} - スキップします。")
        except Exception as e:
            self.log_status(f"❌ 変換失敗: {input_path} - エラー: {e}")

    def convert_xls_to_xlsx(self, input_path):
        try:
            output_path = input_path.replace('.xls', '.xlsx')
            # pandasを使用して.xlsを.xlsxに変換
            excel_data = pd.read_excel(input_path, sheet_name=None)
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                for sheet_name, sheet_data in excel_data.items():
                    sheet_data.to_excel(writer, sheet_name=sheet_name, index=False)
            self.log_status(f"✅ 変換成功: {input_path} -> {output_path}")
            return output_path
        except Exception as e:
            self.log_status(f"❌ .xls を .xlsx に変換できませんでした: {input_path} - エラー: {e}")
            return None

    def process_directory(self, md_converter, input_dir, output_dir):
        file_list = []
        for root_dir, dirs, files in os.walk(input_dir):
            for file in files:
                file_list.append(os.path.join(root_dir, file))

        total_files = len(file_list)
        self.progress['maximum'] = total_files
        processed_files = 0

        for file_path in file_list:
            _, ext = os.path.splitext(file_path)
            if ext.lower() not in self.SUPPORTED_EXTENSIONS:
                self.log_status(f"⚠️ サポートされていない拡張子: {file_path} - スキップします。")
                processed_files += 1
                self.progress['value'] = processed_files
                self.root.update_idletasks()
                continue

            # 出力パスの設定
            rel_path = os.path.relpath(os.path.dirname(file_path), input_dir)
            current_output_dir = os.path.join(output_dir, rel_path)
            if not os.path.exists(current_output_dir):
                os.makedirs(current_output_dir)
            base, _ = os.path.splitext(os.path.basename(file_path))
            output_file_path = os.path.join(current_output_dir, base + '.md')

            self.convert_file(md_converter, file_path, output_file_path)
            processed_files += 1
            self.progress['value'] = processed_files
            self.root.update_idletasks()

        self.progress['value'] = 0  # プログレスバーをリセット

def main():
    root = tk.Tk()
    app = MarkdownConverterGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
