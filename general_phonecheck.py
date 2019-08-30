import os,sys
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import openpyxl
import gen_module as gm
import csv
import re

#result_dir_buttonのコマンド:結果ファイル用のディレクトリゲット
def dir_clicked():
    dir3 = 'C:'
    file3 = filedialog.askdirectory(initialdir=dir3)
    result_dir_box.insert(tk.END,file3)

def process():
    #結果記入ファイルを作成
    result_wb = openpyxl.Workbook()
    #シート名変更
    result_ws = result_wb.active
    result_ws.title = "IP電話チェックリスト"
    #シートに項目を作成
    check_items = ["MACアドレス","IPアドレス","登録","DHCPサーバー","TFTPサーバー1",\
                   "TFTPサーバー2","CUCMサーバー1","CUCMサーバー2","ITLファイル"]
    for i in range(9):
        result_ws.cell(row=1,column=i+1).value = check_items[i]

    #入力されたIPアドレスを取得
    start_ip = start_ip_box.get()
    end_ip = end_ip_box.get()
    #入力されたIPアドレスの第4オクテットを取得
    start_ip_last = re.search("\d{1,3}$",start_ip).group()
    end_ip_last = re.search("\d{1,3}$",end_ip).group()
    #ネットワークセグメントを取得
    net_segment = re.search("(\d{1,3}.){3}\.",start_ip).group()
    #処理開始
    excel_row = 2  #Excelの入力開始位置
    """
    IPアドレス範囲の機器に対して分岐処理。
    pingが通ってかつ、webアクセスができた場合は（つまりIP電話である場合は）
    取得したデータをExcelファイルに書き込む。
    それ以外の場合はエラーメッセージを標準出力する。
    """
    for i in range(int(start_ip_last),int(end_ip_last)+1):
        ip = net_segment+str(i)
        if gm.Ping(ip):
            try:
                mac_addr,result = gm.Phone_Check(ip)
                for j in range(len(result)):
                    result_ws.cell(row=excel_row,column=1).value = mac_addr
                    result_ws.cell(row=excel_row,column=2).value = ip
                    result_ws.cell(row=excel_row,column=3).value = "○"
                    result_ws.cell(row=excel_row,column=j+4).value = result[j]
                excel_row += 1
            except:
                print("もしくはIP電話ではありません")
        else:
            print("疎通がとれません")
    #保存して終了
    result_filename = "/" + res_filename_box.get()
    result_filename.rstrip(".xlsx")
    result_ex = ".xlsx"
    result_count = 0
    filepath = result_dir_box.get() + result_filename + result_ex
    while os.path.exists(filepath):
        int(result_count)
        result_count += 1
        filepath = result_dir_box.get() + result_filename + str(result_count) + result_ex
    result_wb.save(filepath)
    messagebox.showinfo("確認","処理が終了しました。ファイルを確認してください。")

#サイズ要調整
root = tk.Tk()
root.title("IP電話チェックツール")
root.geometry("450x200")
root.resizable(0,0)
#frame1作成
frame1 = tk.Frame(root)
#1行目の部品作成
start_ip_label = tk.Label(frame1,text="開始アドレス")
start_ip_box = tk.Entry(frame1,width=50)
#frame1と1行目を表示
frame1.grid()
start_ip_label.grid(row=0,column=0,padx=5,pady=5)
start_ip_box.grid(row=0,column=1,padx=5,pady=5)
#2行目の部品
end_ip_label = tk.Label(frame1,text="終了アドレス")
end_ip_box = tk.Entry(frame1,width=50)
#2行目を表示
end_ip_label.grid(row=1,column=0,padx=5,pady=5)
end_ip_box.grid(row=1,column=1,padx=5,pady=5)
#3行目の部品
result_dir_label = tk.Label(frame1,text="結果ファイル")
result_dir_box = tk.Entry(frame1,width=50)
result_dir_button = tk.Button(frame1,text="参照",command=dir_clicked)
#3行目を表示
result_dir_label.grid(row=2,column=0,padx=5,pady=5)
result_dir_box.grid(row=2,column=1,padx=5,pady=5)
result_dir_button.grid(row=2,column=2,padx=5,pady=5)
#4行目の部品
res_filename_label = tk.Label(frame1,text="ファイル名")
res_filename_box = tk.Entry(frame1,width=50)
#4行目の表示
res_filename_label.grid(row=3,column=0,padx=5,pady=5)
res_filename_box.grid(row=3,column=1,padx=5,pady=5)
#処理開始ボタン作成
frame2 = tk.Frame(root)
start_button = tk.Button(frame2,text="処理開始",command=process)
#処理開始ボタン表示
frame2.grid()
start_button.grid(row=2,column=1,padx=5,pady=5)

root.mainloop()
