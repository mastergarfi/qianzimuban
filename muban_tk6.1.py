#2021年12月26日22:18:28
import tkinter as tk
from tkinter import *
import tkinter.messagebox
import win32com.client
from win32com.client import Dispatch, constants
import psutil


window=tk.Tk()
window.title('签字单模板6.1')
window.geometry('300x250')

def close_word():
    for proc in psutil.process_iter():
    #关闭所有word进程
        if proc.name() == 'WINWORD.EXE':
            print(proc.pid)
            p = psutil.Process(proc.pid)
            p.terminate()
            print('word关闭')

def mubandy(xm,xb,nl,ch,zyh,zubie,bq):
    print('组别',zubie)
    print('模板打印1')
    zd=''
    ss=''
    if zubie==3:
        daican='麦孚畅清全营养配方粉'
        tongbian='益三联畅达'
        jzys='FCY'
    else:
        daican='永瑞清'
        tongbian = '小麦纤维素颗粒（非比麸）'
        if zubie==1:
            jzys = 'PYN'
        else:
            jzys = 'CY'
    try:
        close_word()
    except:
        print('无法结束word进程')
    # print('print_all',print_all)

    w = win32com.client.Dispatch('Word.Application')

    w.Visible = 0
    w.DisplayAlerts = 0
    w.ScreenUpdating = 0


    doc = w.Documents.Open(r'C:\\1\\1.docx')


    # doc.Content.Find.Execute(FindText=u'abcd', ReplaceWith=u'1234', Replace=2)
    w.Selection.Find.Execute('x_m', False, False, False, False, False, True, 1, True, xm, 2)
    w.Selection.Find.Execute('x_b', False, False, False, False, False, True, 1, True, xb, 2)
    w.Selection.Find.Execute('n_l', False, False, False, False, False, True, 1, True, nl, 2)
    w.Selection.Find.Execute('c_h', False, False, False, False, False, True, 1, True, ch, 2)
    w.Selection.Find.Execute('z_y_h', False, False, False, False, False, True, 1, True, zyh, 2)
    w.Selection.Find.Execute('dai_can', False, False, False, False, False, True, 1, True, daican, 2)
    w.Selection.Find.Execute('b_q', False, False, False, False, False, True, 1, True, bq, 2)
    w.Selection.Find.Execute('tong_bian', False, False, False, False, False, True, 1, True, tongbian, 2)
    w.Selection.Find.Execute('jzys', False, False, False, False, False, True, 1, True, jzys, 2)

    doc.PrintOut()
    w.Documents.Close(SaveChanges=0)
    w.Quit()

def mubandy2(xm,xb,nl,ch,zyh,zubie,bq):
    print('组别',zubie)
    print('模板打印2')
    zd=''
    ss=''
    if zubie==3:
        daican='麦孚畅清全营养配方粉'
        tongbian='益三联畅达'
        jzys='方臣阳'
    else:
        daican='永瑞清'
        tongbian = '小麦纤维素颗粒（非比麸）'
        if zubie==1:
            jzys = '裴艳妮'
        else:
            jzys = '陈勇'

    try:
        close_word()
    except:
        print('无法结束word进程')
    # print('print_all',print_all)

    w = win32com.client.Dispatch('Word.Application')

    w.Visible = 1
    w.DisplayAlerts = 0
    w.ScreenUpdating = 1


    doc = w.Documents.Open(r'C:\\1\\1.docx')


    # doc.Content.Find.Execute(FindText=u'abcd', ReplaceWith=u'1234', Replace=2)
    w.Selection.Find.Execute('x_m', False, False, False, False, False, True, 1, True, xm, 2)
    w.Selection.Find.Execute('x_b', False, False, False, False, False, True, 1, True, xb, 2)
    w.Selection.Find.Execute('n_l', False, False, False, False, False, True, 1, True, nl, 2)
    w.Selection.Find.Execute('c_h', False, False, False, False, False, True, 1, True, ch, 2)
    w.Selection.Find.Execute('z_y_h', False, False, False, False, False, True, 1, True, zyh, 2)
    w.Selection.Find.Execute('dai_can', False, False, False, False, False, True, 1, True, daican, 2)
    w.Selection.Find.Execute('b_q', False, False, False, False, False, True, 1, True, bq, 2)
    w.Selection.Find.Execute('tong_bian', False, False, False, False, False, True, 1, True, tongbian, 2)
    w.Selection.Find.Execute('jzys', False, False, False, False, False, True, 1, True, jzys, 2)



def erase():
    e_xm.delete(0,'end')
    e_xb.delete(0,'end')
    e_nl.delete(0,'end')
    e_ch.delete(0,'end')
    e_zyh.delete(0,'end')
    e_copy.delete(0,'end')
    # e_zd.delete(0,'end')

def get_info():

    l_xm = e_xm.get()
    l_xb = e_xb.get()
    l_nl = e_nl.get()
    l_ch = e_ch.get()
    l_zyh = e_zyh.get()
    l_bq=e_bq.get()
    zubie = v.get()
    '''
    t.insert('insert',l_xm)
    t.insert('insert', l_xb)
    t.insert('insert', l_nl)
    t.insert('insert', l_ch)
    t.insert('insert', l_zyh)
    t.insert('insert', l_zd)
    '''
    #window.quit()

    mubandy(l_xm, l_xb, l_nl, l_ch, l_zyh, zubie,l_bq)


def get_info2():

    l_xm = e_xm.get()
    l_xb = e_xb.get()
    l_nl = e_nl.get()
    l_ch = e_ch.get()
    l_zyh = e_zyh.get()
    l_bq=e_bq.get()
    zubie = v.get()
    '''
    t.insert('insert',l_xm)
    t.insert('insert', l_xb)
    t.insert('insert', l_nl)
    t.insert('insert', l_ch)
    t.insert('insert', l_zyh)
    t.insert('insert', l_zd)
    '''
    #window.quit()

    mubandy2(l_xm, l_xb, l_nl, l_ch, l_zyh, zubie,l_bq)


def fill():
    #print('1')
    info=e_copy.get()
    print(info)

    s = info.find(':')
    end = info.find(' ', s + 1)
    zyh = info[s + 1:end]
    print(zyh)

    s = info.find('床')
    end = info.find(' ', s + 1)
    ch = info[s -2:s]
    print(ch)

    s = info.find('床')
    end = info.find('\n', s + 1)
    xm = info[s + 1:end ]
    print(xm)

    s = info.find('岁')
    end = info.find(' ', s - 3)
    nl = info[end+1:s]
    print(nl)

    s = info.find('床')
    end = info.find('\n', s + 1)
    xb = info[end+1:end+2]
    print(xb)


    e_xm.delete(0, 'end')
    e_xm.insert(0,xm)

    e_xb.delete(0,'end')
    e_xb.insert(0,xb)

    e_nl.delete(0,'end')
    e_nl.insert(0,nl)

    e_ch.delete(0,'end')
    e_ch.insert(0,ch)

    e_zyh.delete(0,'end')
    e_zyh.insert(0,zyh)



var=tk.StringVar()

l_xm=tk.Label(window,text='姓名')
l_xm.place(x=0,y=0)
e_xm=tk.Entry(window)
e_xm.place(x=60,y=0)

l_xb=tk.Label(window,text='性别')
l_xb.place(x=0,y=30)
e_xb=tk.Entry(window)
e_xb.place(x=60,y=30)

l_nl=tk.Label(window,text='年龄')
l_nl.place(x=0,y=60)
e_nl=tk.Entry(window)
e_nl.place(x=60,y=60)

l_bq=tk.Label(window,text='病区')
l_bq.place(x=0,y=90)
e=StringVar()
e_bq=tk.Entry(window,textvariable=e)
e.set('14')
e_bq.place(x=60,y=90)

l_ch=tk.Label(window,text='床号')
l_ch.place(x=0,y=120)
e_ch=tk.Entry(window)
e_ch.place(x=60,y=120)

l_zyh=tk.Label(window,text='住院号')
l_zyh.place(x=0,y=150)
e_zyh=tk.Entry(window)
e_zyh.place(x=60,y=150)

l_copy=tk.Label(window,text='粘贴框')
l_copy.place(x=0,y=180)
e_copy=tk.Entry(window)
e_copy.place(x=60,y=180)

v = tk.IntVar()
#Radiobutto是单选框，只能选中一个
#选中哪一个按钮，会把value的值赋给v
a1 = tk.Radiobutton(window,text="杨组",variable=v,value=1)
a1.place(x=0,y=210)

a2 = tk.Radiobutton(window,text="罗组",variable=v,value=2)
a2.place(x=60,y=210)

a3 = tk.Radiobutton(window,text="郑组",variable=v,value=3)
a3.place(x=120,y=210)
v.set(1)
print(v)

# CheckVar1 = IntVar()
# CheckVar2 = IntVar()
# C1 = Checkbutton(window, text = "RUNOOB", variable = CheckVar1, onvalue = 1, offvalue = 0)
# C2 = Checkbutton(window, text = "GOOGLE", variable = CheckVar2, onvalue = 1, offvalue = 0)
# C1.place(x=0,y=210)
# CheckVar2.set(1)
# C2.place(x=60,y=210)
# print(CheckVar1,CheckVar2)
# vvv=CheckVar1.get()
# print(vvv)


# list1 = ['李白', '杜甫', '李清照', '唐伯虎', '王昭君', '西施']
# v = []
#
# for i in range(len(list1)):
#     v.append(IntVar())
#     check = Checkbutton(window, text=list1[i], variable=v[-1])
#     check.place(x=i*60,y=210)





#显示v的值
# l = tk.Label(window,textvariable=v)
# l.place(x=120,y=180)



# l_zd=tk.Label(window,text='诊断')
# l_zd.place(x=0,y=180)
# e_zd=tk.Entry(window)
# e_zd.place(x=60,y=180)

b0=tk.Button(window,text='填写',command=fill)
b0.pack(anchor=SE)

b=tk.Button(window,text='打印全部',command=get_info)
b.pack(anchor=SE)

b1=tk.Button(window,text='选择打印',command=get_info2)
b1.pack(anchor=SE)

b2=tk.Button(window,text='清空',command=erase)
b2.pack(anchor=SE)



#t=tk.Text(window)
#t.pack(side='bottom')




window.mainloop()
