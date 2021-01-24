import tkinter
from tkinter import ttk
import cx_Oracle as oea
from tkinter import *
# root = Tk()
import tkinter.messagebox as messagebox
db_port = '1521'
ser_name = 'pora12c1.lecent.domain'
msg=[]
bill_number=''
class link_oracle:
    f = open("./oracle_link.txt","r",encoding='UTF-8')
    lines = f.readlines()
    for line in lines:
        line = line.rstrip("\n")
        msg.append(line)

win = tkinter.Tk()
win.title("老资管常见数据问题处理脚本工具(oracle)")
win.geometry("900x700")
#说明
w1 = tkinter.Label(win, text="以下功能直接点击后会自动进行检查，未返回结果之前请勿关闭！！",font=("隶书",9), fg="red")
w1.place(x=250, y=5, anchor='nw')

#连接oracle
def go(sql):  #处理事件，*args表示可变参数
    print(comboxlist.get().split(',')) #打印选中的值
    #获取IP,用户及密码连接Oracle
    status=0
    IP=comboxlist.get().split(',')[0]
    user=comboxlist.get().split(',')[1]
    psd=comboxlist.get().split(',')[2]
    print('ip：'+IP+'用户：'+user+'密码：'+psd)
    conn = oea.connect(user, psd, IP + ':' + '1521' + '/' + 'pora12c1.lecent.domain')
    #conn = oea.connect(user, psd, IP + ':' + '1521' + '/' + 'oracle.lecent.domain')#本地
    link_ora = conn
    cur = conn.cursor()  # 定义连接对象
    link_cur=cur
    # sql = "select * from base_offender_info where 1=1"
    result = cur.execute(sql).fetchall()
    title = [i[0] for i in cur.description]
    # print(title)
    aletr_msg = tkinter.Label(win, text="数据连接成功...",font=("隶书",11), fg="green")
    aletr_msg.place(x=460, y=40, anchor='nw')
    cur.close()  # 关闭cursor对象
    # conn.commit()
    conn.close()  # 关闭数据库链接
    return result

#查询上下账状态不正常的
def prisoner_money_check():
    sql = "select t.deposit_id as 单号,t.bank_card_number as 卡号,t.status as 状态 from capital_deposit_item t \
    where t.status in (0,2,9) union all select cwi.withdrawl_id as 单号,cwi.bank_card_number as 卡号,cwi.status as 状态 \
    from capital_withdrawl_item cwi where cwi.status in (0,2,9)"
    results = go(sql)
    print(results)
    return_msg.delete("1.0",tkinter.END)
    if len(results) > 0:
        return_msg.insert('insert',results)
        # return_msg.pack(fill=tkinter.Y, side=tkinter.BOTTOM)
    else:
        return_msg.insert('insert', '暂无结果！')


def create_newwin():
    top = Toplevel()
    top.title('上下账处理')
    top.geometry("400x300")
    ts = tkinter.Label(top, text="处理银行交易成功但是状态为“处理中”的情况！！", font=("隶书", 10),fg="red")
    ts.place(x=1, y=20, anchor='nw')
    tl = tkinter.Label(top, text="输入单号:", font=("隶书", 10), fg="green")
    tl.place(x=5, y=60, anchor='nw')

    v1 = StringVar()
    e1 = Entry(top, textvariable=v1, width=40)
    bill_number = v1.get()
    e1.grid(row=1, column=0, padx=10, pady=10)
    e1.place(x=10, y=80)

    # 上下账处理处理中的
    def prisoner_money_status():
        print(v1.get())
        moneyin_upsql = "update capital_deposit cd set cd.status = 2 where cd.deposit_id='"+v1.get()+"'"+";"+"commit;"
        results = go(moneyin_upsql)

        print(results)
    Button(top, text='提交处理',command=prisoner_money_status).place(x=310, y=80)













comvalue=tkinter.StringVar()#窗体自带的文本，新建一个值
comboxlist=ttk.Combobox(win,width=45, height=20,font=1,textvariable=comvalue) #初始化
comboxlist["values"]=msg
comboxlist.current(0) #默认选择第一个
# comboxlist.bind("<<ComboboxSelected>>",go) #绑定事件,(下拉列表框被选中时，绑定go()函数)
comboxlist.place(x=10,y=40,anchor='nw')

# link_btn=tkinter.Button(win,text="确定", font=("隶书",10),command=go,width=6, height=1, fg="green")
# link_btn.place(x=400, y=40, anchor='nw')
# comboxlist.pack()

#功能----上下账模块，接见款
w2 = tkinter.Label(win, text="1、处理中检测：状态为处理中的或者异常的；2、同步明细：处理银行交易成功且未同步的明细的；3、状态修改：处理银行未交易的单子为失败",font=("隶书",9), fg="red")
w2.place(x=10, y=130, anchor='nw')


button1 = tkinter.Button(win, text="上下账单检测", font=("隶书",10), command=prisoner_money_check,width=12, height=2, fg="green")
button1.place(x=10, y=90, anchor='nw')
# button1.pack()

button2 = tkinter.Button(win, text="上下账处理", font=("隶书",10), command=create_newwin,width=10, height=2, fg="green")
button2.place(x=150, y=90, anchor='nw')

button3 = tkinter.Button(win, text="上下账处理中状态修改", font=("隶书",10), width=18, height=2, fg="green")
button3.place(x=300, y=90, anchor='nw')

button4 = tkinter.Button(win, text="接见款处理中", font=("隶书",10), width=18, height=2, fg="green")
button4.place(x=450, y=90, anchor='nw')

button5 = tkinter.Button(win, text="接见款处理中状态修改", font=("隶书",10), width=18, height=2, fg="green")
button5.place(x=600, y=90, anchor='nw')

b1 = tkinter.Label(win, text="返回信息：",font=("隶书",10), fg="green")
b1.place(x=5, y=200, anchor='nw')

return_msg=tkinter.Text(win,width=100,height=10)
return_msg.place(x=80, y=150, anchor='nw')


#功能---服务器模块
button6 = tkinter.Button(win, text="服务器空间查询", font=("隶书",10), width=13, height=2, fg="green")
button6.place(x=10, y=300, anchor='nw')

button7 = tkinter.Button(win, text="tomcat日志清除", font=("隶书",10), width=14, height=2, fg="green")
button7.place(x=130, y=300, anchor='nw')

#数据重复
button8 = tkinter.Button(win, text="银行流水重复删除", font=("隶书",10), width=15, height=2, fg="green")
button8.place(x=260, y=300, anchor='nw')

b2 = tkinter.Label(win, text="返回信息：",font=("隶书",10), fg="green")
b2.place(x=20, y=400, anchor='nw')


button2 = tkinter.Button(win, text="导出生成Excel",font=("隶书",15), width=15, height=3, fg="green")
button2.place(x=300, y=600, anchor='nw')

#连接oracle



win.mainloop()