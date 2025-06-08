import tkinter as tk #GUI構築ライブラリ
from tkinter import filedialog, messagebox  #ファイル選択ダイアログ
import openpyxl as px #pythonで.xlsxを開く追加ライブラリ（インストール済み）

#グローバル変数としてワークブックを保持する
wb = None


def open_file():
    global wb
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])

    if file_path:
        wb = px.load_workbook(file_path)
        print(f"ファイルを開きました: {file_path}")

        # 例: 最初のシート名を表示
        sheet_names = wb.sheetnames
        print(f"シート一覧: {sheet_names}")
        messagebox.showinfo("ファイル読み込み完了", "ファイルを正常に読み込みました^^")

def save_file():
    global wb
    if wb is None:
        messagebox.showwarning("エラー", "まずExcelファイル(*.xlsxまたは*.xls)を開いてください😡")
        return
    
    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")], title = "名前を付けて保存")

    if save_path:
        wb.save(save_path)
        messagebox.showinfo("保存完了", f"Excelファイルを保存しました:\n{save_path}")
        print(f"ファイルを保存しました: {save_path}")
    

# GUIウィンドウ作成
root = tk.Tk()
root.title("Excelデータ変換")

#概要を表示する
description = tk.Label(root, text="Excelファイルを読み込み、単位変換とグラフ描画を行います(´・ω・｀)", font=("Helvetica", 9))
description.pack(padx = 10, pady =(10, 5))

#「ファイルを開く」ボタン
open_button = tk.Button(root, text="Excelファイルを開く", command=open_file)
open_button.pack(padx=10, pady=10)

#「名前をつけて保存」ボタン
save_button = tk.Button(root, text="保存先を指定して保存", command=save_file)
save_button.pack(padx=20, pady=10)

#「アプリを終了する」ボタン
quit_button = tk.Button(text = "アプリを終了する", command = root.destroy)
quit_button.pack(padx=30, pady=10)



root.mainloop()