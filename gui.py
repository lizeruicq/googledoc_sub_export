import os.path
import tkinter as tk
from tkinter.messagebox import showinfo
import windnd
import googlesheet
import threading
import cgsrtpy3
from tkinter import ttk


class basedesk():
    def __init__(self, master):
        self.window = master
        self.window.config()
        self.window.title('video toolkit')
        self.window.geometry('400x600')
        initface(self.window)



class initface():
    def __init__(self,master):
        self.master = master
        self.face = tk.Frame(self.master)
        self.face.pack()
        topbarFrm(self.face)


class topbarFrm():
    def __init__(self,master):
        self.master = master
        self.face = tk.Frame(self.master,bg ="green")
        self.face.pack(fill="x")
        self.tab = tab1Frm(self.master)
        tk.Button(self.face, text='谷歌字幕导出',font=('Microsoft YaHei', 12),command=self.changetab1).pack(side ="left")
        tk.Button(self.face, text="视频字幕压制",font=('Microsoft YaHei', 12),command=self.changetab2).pack(side ="left")
        tk.Button(self.face, text='提取字幕文本',font=('Microsoft YaHei', 12),command=self.changetab3).pack(side ="left")
        tk.Button(self.face, text='srt断行', font=('Microsoft YaHei', 12)).pack(side="left")

    def changetab1(self):
        self.tab.face.destroy()
        self.tab = tab1Frm(self.master)
    def changetab2(self):
        self.tab.face.destroy()
        self.tab = tab2Frm(self.master)
    def changetab3(self):
        self.tab.face.destroy()
        self.tab = tab3Frm(self.master)


class tab1Frm():
    def __init__(self, master):
        self.master = master
        self.face = tk.Frame(self.master)
        self.face.pack(expand = "yes",fill = "both")
        tk.Label(self.face,text="时间轴表格(xls)",font=('Microsoft YaHei', 12)).grid(column=0, row=0)
        self.s1= tk.Entry(self.face)
        self.s1.grid(column=1, row=0)
        tk.Label(self.face, text="google文档名称", font=('Microsoft YaHei', 12)).grid(column=0, row=1)
        self.s2=tk.Entry(self.face)
        self.s2.grid(column=1, row=1)
        tk.Label(self.face, text="多语言字幕来源表", font=('Microsoft YaHei', 12)).grid(column=0, row=2)
        self.s3=tk.Entry(self.face)
        self.s3.grid(column=1, row=2)
        tk.Label(self.face, text="谷歌临时存储表格", font=('Microsoft YaHei', 12)).grid(column=0, row=3)
        self.s4=tk.Entry(self.face)
        self.s4.grid(column=1, row=3)
        tk.Label(self.face, text="谷歌字幕表格起点", font=('Microsoft YaHei', 12)).grid(column=0, row=4)
        self.s5=tk.Entry(self.face)
        self.s5.grid(column=1, row=4)
        tk.Label(self.face, text="谷歌字幕表格终点", font=('Microsoft YaHei', 12)).grid(column=0, row=5)
        self.s6=tk.Entry(self.face)
        self.s6.grid(column=1, row=5)
        tk.Label(self.face, text="导出文件名称", font=('Microsoft YaHei', 12)).grid(column=0, row=6)
        self.s7=tk.Entry(self.face)
        self.s7.grid(column=1, row=6)
        tk.Button(self.face, text="执行", font=('Microsoft YaHei', 12),command=self.thread_it).grid(column=1, row=7)
        self.s1.insert("insert", "直接拖时间轴文件进来")
        self.s2.insert("insert","例:崩3内容team本地化")
        self.s3.insert("insert", "例:游戏外视频")
        self.s4.insert("insert", "例:带时间字幕（自动）")
        self.s5.insert("insert", "例:D2949")
        self.s6.insert("insert", "例:K2971")
        self.s7.insert("insert", "自己取名字")

        self.inputlist = [self.s1,self.s2,self.s3,self.s4,self.s5,self.s6,self.s7]
        def dragged_files(files):
            msg = '\n'.join((item.decode('gbk') for item in files))
            print(msg)
            self.s1.delete(0, "end")
            self.s1.insert("insert", msg)
        windnd.hook_dropfiles(self.face, func=dragged_files)

    def readpara(self):
        self.paralist = []
        for i in self.inputlist:
            self.paralist.append(i.get())
        cgsrtpy3.OutputPath = os.path.dirname(self.s1.get())+ "\\"
        googlesheet.findLocalSheet(self.paralist)

    def thread_it(self):

        t = threading.Thread(target=self.readpara)
        t.setDaemon(True)
        t.start()


class tab2Frm():
    def __init__(self, master):
        self.paralist = []
        self.master = master
        self.face = tk.Frame(self.master)
        self.face.pack(expand="yes", fill="both")
        self.t0 = tab2group(self.face, 0)
        self.t1 = tab2group(self.face, 1)
        self.t2 = tab2group(self.face, 2)
        self.t3 = tab2group(self.face, 3)
        self.t4 = tab2group(self.face, 4)

        tk.Button(self.face, text="执行", font=('Microsoft YaHei', 12),command = self.readpara).pack()

    def readpara(self):
        self.paralist.clear()
        self.paralist.append([self.t0.v.get(),self.t0.c.get(),self.t0.f.get()])
        print(self.paralist)


class tab2group():
    def __init__(self,master,number):
        self.temp = ''
        self.number = number
        self.master = master
        self.face = tk.Frame(self.master)
        self.face.pack()
        tk.Label(self.face, text="视频", font=('Microsoft YaHei', 12)).grid(column=0, row=self.number)
        self.v = tk.Entry(self.face, width=10)
        self.v.grid(column=1, row=self.number)
        tk.Label(self.face, text="字幕", font=('Microsoft YaHei', 12)).grid(column=2, row=self.number)
        self.c = tk.Entry(self.face, width=10)
        self.c.grid(column=3, row=self.number)
        tk.Label(self.face, text="字体", font=('Microsoft YaHei', 12)).grid(column=4, row=self.number)
        self.f = ttk.Combobox(self.face, width=10,textvariable=tk.StringVar(), value=('Noto Sans', 'TH Sarabun New', 'Noto Sans KR', 'Noto Sans TC'))
        self.f.grid(column=5, row=self.number)
        windnd.hook_dropfiles(self.v, func=self.dragged_files)
        windnd.hook_dropfiles(self.c, func=self.dragged_files)

    def dragged_files(self,files):
        msg = '\n'.join((item.decode('gbk') for item in files))
        self.temp = msg
        if self.temp.split('.')[-1] == "mp4":
            self.v.insert("insert",self.temp)
        if self.temp.split('.')[-1] == "srt":
            self.c.insert("insert",self.temp)


class tab3Frm():
    def __init__(self, master):
        self.master = master
        self.face = tk.Frame(self.master)
        self.face.pack(expand="yes", fill="both")
        tk.Label(self.face, text="还在施工中...再等等等等一段时间", font=('Microsoft YaHei', 12)).grid(column=0, row=0)

if __name__ == '__main__':
    window = tk.Tk()
    basedesk(window)
    window.mainloop()