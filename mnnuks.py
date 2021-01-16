from tkinter import *
from tkinter import ttk
import time
import pandas as pd
import random
import os
import  pymysql
from dbfread import DBF
import xlwt # 貌似不支持excel2007的xlsx格式
import xlrd
import socket

root= Tk()
root.title('闽南师范大学非计算机专业考试系统')
root.geometry('800x600+300+100') # 这里的乘号不是 * ，而是小写英文字母 x#小写x代表乘号800x600为窗口大小，+300+100窗口显示位置


#获取本机电脑名
myname = socket.getfqdn(socket.gethostname())
#获取本机ip
myaddr = socket.gethostbyname(myname)


with open("./Database/host.dat") as ho:
    host_read = ho.read()
    # host_read = myaddr[:myaddr.index(".",4)] + ho.read()



#打开数据库连接
db = pymysql.connect(
                host=host_read,  #你自己的数据库名
                port=13759,          #端口
                user='jszx',        #用户
                passwd='zx1375',       #你自己设的密码
                db='mnnucc',          #你的数据库名称
                charset='utf8',     #编码方式
            )
# 使用cursor()方法获取操作游标
cur = db.cursor()



def run0(event):
    global g_zkzh,g_status,g_xsxx    #g_status =1表示是二次   g_xsxx表示学生信息
    g_zkzh=zkzh.get()
    stv.set(zkzh.get())
    sql = "select * from ksxxk where zkzh="+g_zkzh
    cur.execute(sql)  # 执行sql语句
    g_xsxx = cur.fetchall()  # 获取查询的所有记录
    if len(g_xsxx) != 0:
        xm.config(text=g_xsxx[0][1])
        xh.config(text=g_xsxx[0][2])
        if g_xsxx[0][4]==1 or g_xsxx[0][4]==2:
            g_status=1
            lb4.place(relx=0.1, rely=0.6)
            mm.place(relx=0.4, rely=0.6, relwidth=0.3, relheight=0.1)
            mm.focus_set()
            btn1.config(state=NORMAL)
        elif g_xsxx[0][4]==3:
            g_status=3
            lb5.place(relx=0.3, rely=0.6)
        else:
            g_status=0
            btn1.config(state=NORMAL)
            btn1.focus_set()
        btn2.config(state=NORMAL)

def run2():   #重新输入准考证号
    zkzh.delete(0, END)  # 清空输入
    xm.config(text="")
    xh.config(text="")
    lb4.place_forget()
    mm.place_forget()

def mkdir(path):
    path = path.strip()# 去除首位空格
    path = path.rstrip("\\")    # 去除尾部 \ 符号
    # 判断路径是否存在# 存在     True# 不存在   False
    isExists = os.path.exists(path)
    # 判断结果
    if not isExists:
        # 如果不存在则创建目录     # 创建目录操作函数
        os.makedirs(path)
        return True
    else:
        __import__('shutil').rmtree(path)
        os.makedirs(path)
        # print(path + ' 目录已存在')# 如果目录存在则不创建，并提示目录已存在
        return False

def jiem(fp,fpw):  #解密
    filepath=fp
    binfile = open(filepath, 'rb') #打开二进制文件
    size = os.path.getsize(filepath) #获得文件大小
    filepath_w = fpw
    binfile_w = open(filepath_w, 'ab+')  # 追加写入
    data = binfile.read(1) #每次输出一个字节
    data = binfile.read()
    content_w= data
    binfile_w.write(content_w)
    print('content',content_w)
    binfile_w.close()
    binfile.close()

def choti():   #新登录开始分配题目
    g_status=1
    mkdir(g_ksfile)
    cur.execute("select * from ksnum")  # 执行sql语句
    results = dict(cur.fetchall())  # 获取查询的所有记录

    if g_xsxx[0][0][2] == '1':     #信息技术
        os.system("copy .\\chodak.* " +  g_ksfile)
        th= "Op" + "%02d"%(random.randint(1, results["newopr_gp"]))
        cur.execute("select * from kstk where TOPIC_NO=1500" + th[2:])  # 执行sql语句
        op_results = cur.fetchall()
        with open(g_ksfile+'\\op.txt','w') as f:
            f.write(op_results[0][2])
        thw = "Wd"+"%02d"%(random.randint(2,results["xxwd_gp"]))
        th += ", " + thw
        jiem(".\\Alldat\\newdat\\" + thw+ ".docx",g_ksfile+'\\'+thw+ ".docx")
        the= "Ex"+"%02d"%(random.randint(2,results["xxex_gp"]))
        th += ", " + the
        jiem(".\\Alldat\\newdat\\" + the + ".xlsx", g_ksfile + '\\' + the + ".xlsx")
        thp= "Pt"+"%02d"%(random.randint(2,results["xxpt_gp"]))
        th += ", " + thp
        jiem(".\\Alldat\\newdat\\" + thp + ".pptx", g_ksfile + '\\' + thp + ".pptx")


    if g_xsxx[0][0][2] == '2':     #python  未改
        th = "Op"+"%02d"%(random.randint(1,results["xxopr_gp"]))
        thw = "Wd"+"%02d"%(random.randint(1,results["xxwd_gp"]))
        th += ", " + thw
        os.system("copy .\\Alldat\\newdat\\" + thw + ".docx " + g_ksfile)
        the= "Ex"+"%02d"%(random.randint(1,results["xxex_gp"]))
        th += ", " + the
        os.system("copy .\\Alldat\\newdat\\" + the + ".xlsx " + g_ksfile)
        thp= "Pt"+"%02d"%(random.randint(1,results["xxpt_gp"]))
        th += ", " + thp
        os.system("copy .\\Alldat\\newdat\\" + thp + ".pptx " + g_ksfile)

    if g_xsxx[0][0][2] == '6':     #acce 未改
        th = "Op"+"%02d"%(random.randint(1,results["xxopr_gp"]))
        thw = "Wd"+"%02d"%(random.randint(1,results["xxwd_gp"]))
        th += ", " + thw
        os.system("copy .\\Alldat\\newdat\\" + thw + ".docx " + g_ksfile)
        the= "Ex"+"%02d"%(random.randint(1,results["xxex_gp"]))
        th += ", " + the
        os.system("copy .\\Alldat\\newdat\\" + the + ".xlsx " + g_ksfile)
        thp= "Pt"+"%02d"%(random.randint(1,results["xxpt_gp"]))
        th += ", " + thp
        os.system("copy .\\Alldat\\newdat\\" + thp + ".pptx " + g_ksfile)

    sql_update = "update ksxxk set username = '%s',status = '%d',topic = '%s' where zkzh = %s"
    cur.execute(sql_update % ("k:/"+os.environ['USERNAME'], 1, th, g_xsxx[0][0]))  # 像sql语句传递参数
    db.commit()  # 提交
    #拷贝时间文件
    os.system("copy .\\Database\\time.dat " + g_ksfile)
    return

def run1_1():
    global root  # //控制外部界面
    global root1
    global g_star
    root.destroy()
    root1 = Tk()
    root1.title("开始抽题。。。。。。")
    root1.geometry('400x300+300+200')


    image_text = Text(root1, height=16, width=80)
    # image_text.insert('0.0', '\n文本框插入图片')
    photo = PhotoImage(file=g_zkzh[2]+".gif")
    image_text.image_create('1.0', image=photo)
    image_text.grid(row=12, column=1, columnspan=3)
    btn1 = Button(root1, text='开始考试', command=run3, font=('隶书', 16, 'bold'))
    btn1.place(relx=0.3, rely=0.8, relwidth=0.25, relheight=0.15)
    g_star = time.time()
    root1.mainloop()


def run1():
    global g_ksfile   #考生文件夹
    g_ksfile="K:\\"+os.environ["USERNAME"]+"\\"+str(g_xsxx[0][0])   #*******注意修改为k:盘

    if g_status == 0 or mm.get()=="456":  #新考和重抽题
        choti()
        run1_1()
    elif g_status == 1 and mm.get()=="123" and os.environ['USERNAME']==g_xsxx[0][3][3:]:  #二次登录
        run1_1()
    else:
        xm.config(text="")
        xh.config(text="")
        lb4.place_forget()
        mm.place_forget()


def run3():  #考试界面
    root1.destroy()
    root2 = Tk()
    root2.title(g_xsxx[0][0]+"  "+g_xsxx[0][1])
    root2.geometry('800x600')
    with open(g_ksfile + "\\time.dat", 'r') as fr:
        time1 = int(fr.read())
    def gettime():
        time_lb.config(text=str(time1-int((time.time()-g_star))))
        if (int(str(time1-int((time.time()-g_star))))) % 60 == 0:
            with open(g_ksfile+"\\time.dat",'w') as fw:
                fw.write(str(time1-int((time.time()-g_star))))
        if (time1-int((time.time()-g_star)))>0:  #最后应该改为0
            root2.after(1000, gettime)  # 每隔1s调用函数 gettime 自身获取时间
        else:
            print("时间已到")
    time_lb = Label(root2, text='', fg='blue', font=("黑体", 20))
    time_lb.pack()
    gettime()

    def offi():
        jb_btn1.config(state = NORMAL)
        treeview.place_forget()
        chotext.place_forget()
        wintext.place_forget()
        rb_style_a.place_forget()
        rb_style_b.place_forget()
        rb_style_c.place_forget()
        rb_style_d.place_forget()
        write_cho()
        fileB.place(relx=0.66, rely=0.83, relwidth=0.35, relheight=0.05)

        jump_main.post(root2.winfo_pointerxy()[0]-30,root2.winfo_pointerxy()[1]-200)  # post为弹出菜单
    # -----------------------------------------------------------#
    def WordOp():
        print("word")
    def ExcelOp():
        print("exd")
    def PowerPointOp():
        print("ppt")

    jump_main = Menu(root2, tearoff=0)
    # for i in ['word', 'excel', 'ppt']:  # 利用for循环逐一给菜单增添下来菜单
    #     jump_main.add_command(label=i,command=wep)  # label是设置下拉菜单的名称
    jump_main.add_command(label='Word', command=WordOp)
    jump_main.add_command(label='Excel', command=ExcelOp)
    jump_main.add_command(label='PowerPoint', command=PowerPointOp)





    def jiaoj():
        pass


    def win():   #windows操作题
        jb_btn1.config(state = NORMAL)
        treeview.place_forget()
        chotext.place_forget()
        rb_style_a.place_forget()
        rb_style_b.place_forget()
        rb_style_c.place_forget()
        rb_style_d.place_forget()

        write_cho()

        wintext.delete('0.0', END)
        with open(g_ksfile + '\\op.txt', 'r') as f:
            op=f.read()
        wintext.insert('0.0', "考生文件夹："+g_ksfile+"\n\r\n"+op)
        wintext.place(relx=0.15, rely=0.15, relwidth=0.65, relheight=0.5)
        fileB.place(relx=0.66, rely=0.83, relwidth=0.35, relheight=0.05)

    def open_kxfile():#打开考生文件夹
        os.system('start explorer ' + g_ksfile)


    table = DBF(g_ksfile+'\\chodak.DBF', encoding='gbk', char_decode_errors='ignore')
    def cho():
        jb_btn1.config(state = DISABLED)
        wintext.place_forget()
        fileB.place_forget()

        treeview.column("题号", width=40)  # 表示列,不显示
        treeview.column("选项", width=40)
        treeview.heading("题号", text="题号")
        treeview.heading("选项", text="选项")
        def set_cell_value(event):  # 单击进入编辑状态
            def function1():
                rb_style_a.config(relief=SUNKEN)
                rb_style_b.config(relief=FLAT)
                rb_style_c.config(relief=FLAT)
                rb_style_d.config(relief=FLAT)
                treeview.set(item=treeview.selection(), column=1, value="A")

            def function2():
                rb_style_a.config(relief=FLAT)
                rb_style_b.config(relief=SUNKEN)
                rb_style_c.config(relief=FLAT)
                rb_style_d.config(relief=FLAT)
                treeview.set(item=treeview.selection(), column=1, value="B")

            def function3():
                rb_style_a.config(relief=FLAT)
                rb_style_b.config(relief=FLAT)
                rb_style_c.config(relief=SUNKEN)
                rb_style_d.config(relief=FLAT)
                treeview.set(item=treeview.selection(), column=1, value="C")

            def function4():
                rb_style_a.config(relief=FLAT)
                rb_style_b.config(relief=FLAT)
                rb_style_c.config(relief=FLAT)
                rb_style_d.config(relief=SUNKEN)
                treeview.set(item=treeview.selection(), column=1, value="D")

            if len(treeview.identify_row(event.y)) != 0:
                chotext.delete('0.0', END)
                for record in table:
                    if record["TOPIC_BH"] == int(treeview.identify_row(event.y)[-2:], 16):
                        chotext.insert('0.0', record["TOPIC_TXT"])
                        rb_style_a.config(text=record["TOPIC_A"].strip().replace("\r",""), relief=FLAT, command=function1)   #去除头尾空格，替换回车符
                        rb_style_b.config(text=record["TOPIC_B"].strip().replace("\r",""), relief=FLAT, command=function2)
                        rb_style_c.config(text=record["TOPIC_C"].strip().replace("\r",""), relief=FLAT, command=function3)
                        rb_style_d.config(text=record["TOPIC_D"].strip().replace("\r",""), relief=FLAT, command=function4)
                        if treeview.item(treeview.identify_row(event.y), 'values')[1] == "A":
                            rb_style_a.config(relief=SUNKEN)
                        elif treeview.item(treeview.identify_row(event.y), 'values')[1] == "B":
                            rb_style_b.config(relief=SUNKEN)
                        elif treeview.item(treeview.identify_row(event.y), 'values')[1] == "C":
                            rb_style_c.config(relief=SUNKEN)
                        elif treeview.item(treeview.identify_row(event.y), 'values')[1] == "D":
                            rb_style_d.config(relief=SUNKEN)

        treeview.place(relx=0.01, rely=0.05)
        if treeview.selection()==():
            nrows=1
            if os.path.isfile(g_ksfile+"\\cho.dat"):
                book = xlrd.open_workbook(g_ksfile+'\\cho.dat')
                sheet1 = book.sheets()[0]
                nrows = sheet1.nrows
            if nrows > 10:
                for i in range(1,nrows):
                    treeview.insert('',i,values=(sheet1.row_values(i)))
            else:
                for record in table:  # 写入数据
                    treeview.insert('', record["TOPIC_BH"], values=(record["TOPIC_BH"], record["TOPIC_ANS"]))

        treeview.bind('<Button-1>', set_cell_value)  # 双击左键进入编辑

        chotext.place(relx=0.15, rely=0.05, relwidth=0.75, relheight=0.4)
        rb_style_a.place(relx=0.15, rely=0.5)
        rb_style_b.place(relx=0.55, rely=0.5)
        rb_style_c.place(relx=0.15, rely=0.71)
        rb_style_d.place(relx=0.55, rely=0.71)

    def write_cho():
        t = treeview.get_children()
        if len(t) !=0:
            workbook = xlwt.Workbook()
            worksheet = workbook.add_sheet('test')
            worksheet.write(0, 0,"题号")
            worksheet.write(0, 1, "选项")  #第0行，第1列
            for i in t:
                worksheet.write(int(treeview.item(i, 'values')[0]), 0, treeview.item(i, 'values')[0])
                worksheet.write(int(treeview.item(i, 'values')[0]), 1, treeview.item(i, 'values')[1])
            workbook.save(g_ksfile+'\\cho.dat')

    # if g_zkzh[2]=='0':
    #     jb_btn1 = Button(root2, text='选择题', command=cho, font=('隶书', 16, 'bold'))
    #     jb_btn1.place(relx=0.04, rely=0.85, relwidth=0.13, relheight=0.1)
    #     jb_btn2 = Button(root2, text='windows操作题', command=win, font=('隶书', 16, 'bold'))
    #     jb_btn2.place(relx=0.19, rely=0.85, relwidth=0.23, relheight=0.1)
    #     jb_btn3 = Button(root2, text='文字录入', command=win, font=('隶书', 16, 'bold'))
    #     jb_btn3.place(relx=0.44, rely=0.85, relwidth=0.13, relheight=0.1)
    #     jb_btn4 = Button(root2, text='office操作题', command=cho, font=('隶书', 16, 'bold'))
    #     jb_btn4.place(relx=0.59, rely=0.85, relwidth=0.23, relheight=0.1)
    #     jb_btn5 = Button(root2, text='交卷', command=win, font=('隶书', 16, 'bold'))
    #     jb_btn5.place(relx=0.83, rely=0.85, relwidth=0.13, relheight=0.1)

    if g_zkzh[2]=='1':
        jb_btn1 = Button(root2, text='选择题', command= cho, font=('隶书', 18, 'bold'))
        jb_btn1.place(relx=0.05, rely=0.89, relwidth=0.15, relheight=0.1)
        jb_btn2 = Button(root2, text='windows操作题', command=win, font=('隶书', 18, 'bold'))
        jb_btn2.place(relx=0.23, rely=0.89, relwidth=0.25, relheight=0.1)
        jb_btn3 = Button(root2, text='office操作题', command=offi, font=('隶书', 18, 'bold'))
        jb_btn3.place(relx=0.51, rely=0.89, relwidth=0.25, relheight=0.1)
        jb_btn4 = Button(root2, text='交卷', command=jiaoj, font=('隶书', 18, 'bold'))
        jb_btn4.place(relx=0.8, rely=0.89, relwidth=0.15, relheight=0.1)

    if g_zkzh[2]=='2':   #未改
        jb_btn1 = Button(root2, text='选择题', command= cho, font=('隶书', 18, 'bold'))
        jb_btn1.place(relx=0.05, rely=0.85, relwidth=0.15, relheight=0.1)
        jb_btn2 = Button(root2, text='windows操作题', command=win, font=('隶书', 18, 'bold'))
        jb_btn2.place(relx=0.23, rely=0.85, relwidth=0.25, relheight=0.1)
        jb_btn3 = Button(root2, text='office操作题', command=offi, font=('隶书', 18, 'bold'))
        jb_btn3.place(relx=0.51, rely=0.85, relwidth=0.25, relheight=0.1)
        jb_btn4 = Button(root2, text='交卷', command=jiaoj, font=('隶书', 18, 'bold'))
        jb_btn4.place(relx=0.8, rely=0.85, relwidth=0.15, relheight=0.1)

    if g_zkzh[2] == '6':    #未改
        jb_btn1 = Button(root2, text='选择题', command=cho, font=('隶书', 18, 'bold'))
        jb_btn1.place(relx=0.05, rely=0.85, relwidth=0.15, relheight=0.1)
        jb_btn2 = Button(root2, text='windows操作题', command=win, font=('隶书', 18, 'bold'))
        jb_btn2.place(relx=0.23, rely=0.85, relwidth=0.25, relheight=0.1)
        jb_btn3 = Button(root2, text='office操作题', command=offi, font=('隶书', 18, 'bold'))
        jb_btn3.place(relx=0.51, rely=0.85, relwidth=0.25, relheight=0.1)
        jb_btn4 = Button(root2, text='交卷', command=jiaoj, font=('隶书', 18, 'bold'))
        jb_btn4.place(relx=0.8, rely=0.85, relwidth=0.15, relheight=0.1)

    columns = ("题号", "选项")
    treeview = ttk.Treeview(root2, height=20, show="headings", columns=columns)  # 表格
    chotext = Text(root2,height=16, width=45)
    chotext.insert(0.0, "请点击选择题号：")
    # iv_style = IntVar()
    # rb_style_a = Radiobutton(root2, value=1, variable=iv_style, wraplength=300) #设置自动换行wraplength=300
    # rb_style_b = Radiobutton(root2, value=2, variable=iv_style, wraplength=300)
    # rb_style_c = Radiobutton(root2, value=3, variable=iv_style, wraplength=300)
    # rb_style_d = Radiobutton(root2, value=4, variable=iv_style, wraplength=300)
    rb_style_a = Button(root2)
    rb_style_b = Button(root2)
    rb_style_c = Button(root2)
    rb_style_d = Button(root2)

    wintext = Text(root2, height=16, width=45)
    fileB = Button(root2, text='考生文件夹：'+g_ksfile, command= open_kxfile, font=('隶书', 12, 'bold'))


    root2.mainloop()


lb1 = Label(root, text='准考证号', font=('隶书', 32, 'bold'),bg='#d3fbfb')
lb1.place(relx=0.1, rely=0.15)
lb2 = Label(root, text='姓    名', font=('隶书', 32, 'bold'),bg='#d3fbfb')
lb2.place(relx=0.1, rely=0.3)
lb3 = Label(root, text='学    号', font=('隶书', 32, 'bold'),bg='#d3fbfb')
lb3.place(relx=0.1, rely=0.45)
lb4 = Label(root, text='密    码', font=('隶书', 32, 'bold'), bg='#d3fbfb')
mm = Entry(root, font=('隶书', 32, 'bold'),show ='*')
lb5 = Label(root, text='该考生已评分！', font=('隶书', 32, 'bold'), fg='#ff0000', bg='#fffbfb')

zkzh = Entry(root,width=7,font=('隶书', 32, 'bold'),bg='#d3fbfb')
zkzh.place(relx=0.4, rely=0.15, relwidth=0.3, relheight=0.1)
stv = StringVar()
zkzh.config(textvariable=stv)
zkzh.bind('<Key-Return>', run0)

xm = Label(root,font=('隶书', 32, 'bold'))
xm.place(relx=0.4,rely=0.3, relwidth=0.3, relheight=0.1)
xh = Label(root,font=('隶书', 32, 'bold'))
xh.place(relx=0.4,rely=0.45, relwidth=0.3, relheight=0.1)

btn1 = Button(root, text='确  定', command=run1,state=DISABLED,font=('隶书', 18, 'bold'))
btn1.place(relx=0.3, rely=0.85, relwidth=0.15, relheight=0.1)
btn2 = Button(root, text='重  输', command=run2,state=DISABLED,font=('隶书', 18, 'bold'))
btn2.place(relx=0.6, rely=0.85, relwidth=0.15, relheight=0.1)
root.mainloop()
