import tkinter as tk #GUI構築ライブラリ
from tkinter import filedialog #ファイル選択ダイアログ
import openpyxl as px #pythonで.xlsxを開く追加ライブラリ（インストール済み）
# import xlrd as px #pythonで.xlsを開く追加ライブラリ（インストール済み）

def open_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])

    if file_path:
        wb = px.load_workbook(file_path)
        print(f"ファイルを開きました: {file_path}")

        # 例: 最初のシート名を表示
        sheet_names = wb.sheetnames
        print(f"シート一覧: {sheet_names}")

# GUIウィンドウ作成
root = tk.Tk()
root.title("Excelデータ変換アプリ")

#概要を表示する
description = tk.Label(root, text="Excelファイルを読み込み、単位変換とグラフ描画を行います(´・ω・｀)", font=("Helvetica", 9))
description.pack(padx = 10, pady =(10, 5))

#「ファイルを開く」ボタン
open_button = tk.Button(root, text="Excelファイルを開く", command=open_file)
open_button.pack(padx=10, pady=10)

root.mainloop()

#wb = plx.load_workbook("test.xlsx") #wbに指定ワークブックを入れる

#wb.save("../test2.xlsx") #指定ワークブックにファイルパスを与えて保存