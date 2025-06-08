import pandas as pd
import tkinter as tk #GUI構築ライブラリ
from tkinter import filedialog, messagebox  #ファイル選択ダイアログ
import openpyxl as px #pythonで.xlsxを開く追加ライブラリ
import xlrd as pxl #pythonで.xlsを開く追加ライブラリ
import os

#グローバル変数としてワークブックを保持する
wb = None

def read_excel_file(filepath):
    if filepath.endswith(".xls"):
        df = pd.read_excel(filepath, engine="xlrd")
    else:
        df = pd.read_excel(filepath, engine="openpyxl")

    return df

def open_file():
    global wb
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])

    if file_path:
        filename = os.path.basename(file_path)
        file_label.config(text=f"{filename}")

        if file_path.endswith(".xls"):
            try:
                df = pd.read_excel(file_path, engine="xlrd")
                wb = None
                analyze_columns_xls(df)
            except Exception as e:
                messagebox.showerror("エラー", f"ファイルを読み込めませんでした\n{e}")
        else:
            try:
                wb = px.load_workbook(file_path, data_only=True)
                analyze_columns_xlsx(file_path)
            except Exception as e:
                messagebox.showerror("エラー", f"ファイルを読み込めませんでした\n{e}")

    else:
        file_label.config(text="ファイル未選択")
        columns_text.set("")
        return

def analyze_columns_xls(df):
    result_lines = []

    for col in df.columns:
        count = df[col].apply(lambda x: isinstance(x, (int, float))).sum()
        result_lines.append(f"{col}: {count}件の数値データ")

    columns_text.set("\n".join(result_lines))


def analyze_columns_xlsx(filepath):
    try:
        if filepath.endswith(".xls"):
            df = pd.read_excel(filepath, engine="xlrd", header=None)
        else:
            wb = px.load_workbook(filepath, data_only=True)
            sheet = wb.active
            max_col = sheet.max_column
            max_row = sheet.max_row

            result_lines = []

            for col_idx in range(1, max_col +1):
                part1 = sheet.cell(row=1, column=col_idx).value or ""
                part2 = sheet.cell(row=1, column=col_idx).value or ""
                col_name = f"{part1}/{part2}".strip(" /")

                count = 0
                for row in range (2, max_row +1):
                    cell = sheet.cell(row=row, column=col_idx)
                    if isinstance(cell.value, (int, float)):
                        count += 1

                result_lines.append(f"{col_name}: {count}件の数値データ")

            columns_text.set("\n".join(result_lines))
            return
    
        # xlsの場合（pandas使用）
        result_lines = []
        for col in df.columns:
            count = pd.to_numeric(df[col], errors='coerce').notna().sum()
            result_lines.append(f"{col_name}: {count}件の数値データ")
        
        columns_text.set("\n".join(result_lines))

    except Exception as e:
        messagebox.showerror("エラー", f"ファイルを読み込めませんでした\n{e}")


def save_file():
    global wb
    if wb is None:
        messagebox.showwarning("エラー", "まずExcelファイル(*.xlsxまたは*.xls)を開いてくださいヽ(`Д´)ﾉ")
        return
    
    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")], title = "名前を付けて保存")

    if save_path:
        wb.save(save_path)
        messagebox.showinfo("保存完了", f"Excelファイルを保存しました:\n{save_path}")
        print(f"ファイルを保存しました: {save_path}")
    
def quit_app():
    tk.Tk().withdraw()
    res = messagebox.askokcancel('アプリ終了', 'アプリを終了しますか？')

    if res is True:
        root.quit()


# GUIウィンドウ作成
root = tk.Tk()
root.title("Excelデータ変換")

# ウィンドウの幅と高さ
window_width = 600
window_height = 400

# 画面サイズの取得
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

# 中央座標の計算
x = int((screen_width - window_width) / 2)
y = int((screen_height - window_height) / 2)

# ウィンドウの位置を設定
root.geometry(f"{window_width}x{window_height}+{x}+{y}")

#概要を表示する
description = tk.Label(root, text="Excelファイルを読み込み、単位変換とグラフ描画を行います(´・ω・｀)", font=("Helvetica", 9))
description.pack(padx = 10, pady =(10, 5))

#「ファイルを開く」ボタン
open_button = tk.Button(root, text="Excelファイルを開く", width=20, height=3, command=open_file)
open_button.pack(padx=10, pady=10)

#ファイル名表示用ラベル
file_label = tk.Label(root, text="ファイル未選択", fg="blue", font=("Helvetica", 10))
file_label.pack(padx=10, pady=10)

columns_text = tk.StringVar()
tk.Label(root, textvariable=columns_text, justify="left", wraplength=550, anchor="w").pack(pady=10)

#「名前をつけて保存」ボタン
save_button = tk.Button(root, text="保存先を指定して保存", width=20, height=3, command=save_file)
save_button.pack(padx=20, pady=10)

#「アプリを終了する」ボタン
quit_button = tk.Button(text = "アプリを終了する", width=20, height=3, command = quit_app)
quit_button.pack(padx=30, pady=10)

root.mainloop()

