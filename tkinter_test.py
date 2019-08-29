import os,sys
import tkinter as tk
from tkinter import filedialog
#button1のコマンド:csvファイル名をゲット
def csv_clicked():
    ftype1 = [("csvファイル","*.csv")]
    dir1 = 'C:'
    file1 = filedialog.askopenfilename(filetypes=ftype1,initialdir=dir1)
    entrybox1.insert(tk.END,file1)
#button2のコマンド:excelファイル名をゲット
def excel_clicked():
    ftype2 = [("Excelファイル","*.xlsx")]
    dir2 = 'C:'
    file2 = filedialog.askopenfilename(filetypes=ftype2,initialdir=dir2)
    entrybox2.insert(tk.END,file2)
#button3のコマンド:結果ファイル用のディレクトリゲット
def dir_clicked():
    dir3 = 'C:'
    file3 = filedialog.askdirectory(initialdir=dir3)
    entrybox3.insert(tk.END,file3)
#サイズ要調整
root = tk.Tk()
root.title("testkinter")
root.geometry("450x160")
root.resizable(0,0)
#frame1作成
frame1 = tk.Frame(root)
#1行目の部品作成
file1 = ""
label1 = tk.Label(frame1,text="csvファイル")
entrybox1 = tk.Entry(frame1,width=50)
button1 = tk.Button(frame1,text="参照",command=csv_clicked)
#frame1と1行目を表示
frame1.grid()
label1.grid(row=0,column=0,padx=5,pady=5)
entrybox1.grid(row=0,column=1,padx=5,pady=5)
button1.grid(row=0,column=2,padx=5,pady=5)
#2行目の部品
label2 = tk.Label(frame1,text="Excelファイル")
entrybox2 = tk.Entry(frame1,width=50)
button2 = tk.Button(frame1,text="参照",command=excel_clicked)
#2行目を表示
label2.grid(row=1,column=0,padx=5,pady=5)
entrybox2.grid(row=1,column=1,padx=5,pady=5)
button2.grid(row=1,column=2,padx=5,pady=5)
#3行目の部品
label3 = tk.Label(frame1,text="結果ファイル")
entrybox3 = tk.Entry(frame1,width=50)
button3 = tk.Button(frame1,text="参照",command=dir_clicked)
#3行目を表示
label3.grid(row=2,column=0,padx=5,pady=5)
entrybox3.grid(row=2,column=1,padx=5,pady=5)
button3.grid(row=2,column=2,padx=5,pady=5)
#処理開始ボタン作成
frame2 = tk.Frame(root)
start_button = tk.Button(frame2,text="処理開始",command=)
#処理開始ボタン表示
frame2.grid()
start_button.grid(row=2,column=1,padx=5,pady=5)

root.mainloop()
