import tkinter
import cx_Oracle as oea
from tkinter import *
import tkinter.filedialog
from tkinter import messagebox  # 弹出提示框

win = tkinter.Tk()
win.title("退货时间修改工具(oracle)")
win.geometry("500x300")
win.resizable(0, 0)  # 锁定窗口
#说明
w1 = tkinter.Label(win, text="以下功能直接点击后会自动进行检查，未返回结果之前请勿关闭！！",font=("隶书",9), fg="red")
w1.place(x=80, y=5, anchor='nw')

#连接oracle
def go(sql):  #处理事件，*args表示可变参数
    # conn = oea.connect('LECENT', 'lecenTkm3001', '10.9.101.180' + ':' + '1521' + '/' + 'pora12c1.lecent.domain')
    conn = oea.connect('lecent_ank', 'lecent_ank', '192.168.1.25' + ':' + '1521' + '/' + 'pora12c1.lecent.domain')
    print(conn)
    aletr_msg = tkinter.Label(win, text="数据连接成功...", font=("隶书", 11), fg="green")
    aletr_msg.place(x=650, y=40, anchor='nw')
    cur = conn.cursor()  # 定义连接对象
    if sql.split()[0] == 'select':
        result = cur.execute(sql).fetchall()
        cur.close()  # 关闭cursor对象
        conn.close()  # 关闭数据库链接
        return result

    if (sql.split()[0] == 'update' or sql.split()[0] == 'delete' or sql.split()[0] == 'create' or sql.split()[0] == 'drop') :
        cur.execute(sql)
        effectRow = cur.rowcount
        conn.commit()  # 执行update操作时需要写这个，否则就会更新不成功
        cur.close()  # 关闭cursor对象
        conn.close()  # 关闭数据库链接
        return effectRow


def sale_return():
    if  num.get().isalnum() == True:
        if int(num.get()) <= 30 and int(num.get()) > 0:
            msg_show.set('')
            print(num.get())
            update_sql = "update base_system_configuration set sale_return="+num.get()
            result = go(update_sql)
            msg_show_suc.set("修改成功！！")
            messagebox.showinfo('提示',"退货时间已经修改为"+num.get()+"天")
        else:
            msg_show_suc.set('')
            msg_show.set("天数介于[1-30]之间！！")

w2 = tkinter.Label(win, text="输入退货天数",font=("隶书",9), fg="red")
w2.place(x=10, y=50, anchor='nw')

msg_show = StringVar()
alert = tkinter.Label(win, textvariable=msg_show,font=("隶书",9), fg="green")
alert.place(x=260, y=120, anchor='nw')

msg_show_suc = StringVar()
alert_suc = tkinter.Label(win, textvariable=msg_show_suc,font=("隶书",9), fg="green")
alert_suc.place(x=260, y=80, anchor='nw')

num = StringVar()
sale_num = Entry(win, textvariable=num, width=30)
sale_num.grid(row=1, column=0, padx=10, pady=10)
sale_num.place(x=10, y=80)

button1 = tkinter.Button(win, text="退货天数修改", font=("隶书",10), command=sale_return,width=12, height=2, fg="green")
button1.place(x=10, y=120, anchor='nw')


win.mainloop()