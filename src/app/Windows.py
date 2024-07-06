# -*- coding: utf-8 -*-

'''
檔案說明：程式視窗等元件
Writer：Qian
'''

import tkinter as tk

win = tk.Tk()

#標題
win.title("BD Data Processing")

#大小
win.geometry("1280x720+50+50")
win.minsize(width = 1280, height = 720)
win.config(bg="#272727")
# 272727、D4AA7D、EFD09E、D2D8B3、90A9B7
a = tk.Label(win, text="AAA", background="#D4AA7D")
a.place(relx=0.025,rely=0.03)

funbtn = tk.Button(win,\
                    text="btn",\
                    width=20,\
                    height=5
                    )
funbtn.place(relx=0.015,rely=0.1)

win.iconbitmap("./docs/design/data-collection.ico")
win.mainloop()