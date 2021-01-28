import tkinter
from tkinter import ttk
import cx_Oracle as oea
from tkinter import *
import  time
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
    # print(comboxlist.get().split(',')) #打印选中的值
    #获取IP,用户及密码连接Oracle
    status=0
    IP=comboxlist.get().split(',')[0]
    user=comboxlist.get().split(',')[1]
    psd=comboxlist.get().split(',')[2]
    # print('ip：'+IP+'用户：'+user+'密码：'+psd)
    conn = oea.connect(user, psd, IP + ':' + '1521' + '/' + 'pora12c1.lecent.domain')
    aletr_msg = tkinter.Label(win, text="数据连接成功...", font=("隶书", 11), fg="green")
    aletr_msg.place(x=600, y=40, anchor='nw')
    #conn = oea.connect(user, psd, IP + ':' + '1521' + '/' + 'oracle.lecent.domain') #本地测试
    cur = conn.cursor()  # 定义连接对象
    # sql = "select * from base_offender_info where 1=1"
    if sql.split()[0] == 'select':
        result = cur.execute(sql).fetchall()
        #title = [i[0] for i in cur.description]
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

#连接服务器
def link_linux_server():
    print('11111')



#查询上下账状态不正常的
def prisoner_money_check():
    sql = "select * from (select t.deposit_id as 单号,t.bank_card_number as 卡号,t.status as 状态 from capital_deposit_item t \
    where t.status in (0,2,9) union all select cwi.withdrawl_id as 单号,cwi.bank_card_number as 卡号,cwi.status as 状态 \
    from capital_withdrawl_item cwi where cwi.status in (0,2,9)) order by 状态 desc "
    results = go(sql)

    bill_num.delete("1.0", tkinter.END)
    sql_bill = "select 单号 from (select t.deposit_id as 单号,t.bank_card_number as 卡号,t.status as 状态 from capital_deposit_item t \
    where t.status in (0,2,9) union all select cwi.withdrawl_id as 单号,cwi.bank_card_number as 卡号,cwi.status as 状态 \
    from capital_withdrawl_item cwi where cwi.status in (0,2,9)) group by 单号  "
    bill_results = go(sql_bill)
    bill_num.insert('insert',bill_results)

    x = return_msg.get_children()
    for item in x:
        return_msg.delete(item)

    if len(results) > 0:
        # return_msg.insert('insert',results)
        j = 0
        for i in results:
            return_msg.insert('', 'end', values=results[j])
            j=j+1

        # bill_num.pack(fill=tkinter.Y, side=tkinter.BOTTOM)
    else:
        bill_num.insert('insert', '暂无结果！')


def create_newwin():
    top = Toplevel()
    top.title('上下账处理')
    top.geometry("600x300")
    ts = tkinter.Label(top, text="处理银行交易成功但是状态为“处理中”且上下账单明细未同步的情况！！", font=("隶书", 10),fg="red")
    ts.place(x=1, y=20, anchor='nw')
    tl = tkinter.Label(top, text="输入单号:", font=("隶书", 10), fg="green")
    tl.place(x=5, y=60, anchor='nw')

    v1 = StringVar()
    e1 = Entry(top, textvariable=v1, width=40)
    # bill_number = v1.get()
    e1.grid(row=1, column=0, padx=10, pady=10)
    e1.place(x=10, y=80)

    # 上下账处理处理中的
    def prisoner_money_status():
        str = StringVar()
        alert_msg = tkinter.Label(top, textvariable=str, font=("隶书", 10), fg="green")
        alert_msg.place(x=1, y=150, anchor='nw')
        if v1.get()[0:2] == 'CD':
            deposit_upsql = "update capital_deposit cd set cd.status = 2 where cd.deposit_id='"+v1.get()+"'"
            results_deposit = go(deposit_upsql)
            if results_deposit == 0 :
                str.set('未找到上账单！')
            else:
                #查询中间表是否存在，或者是否存在多条明细，无中间表则返回提示未生成中间表，有多条中间表则删除vchr_st为空的中间表
                proxy_count = "select count(*) from bank_proxy_payment_received bpr where bpr.cust_bill_number='"+v1.get()+"'"
                total_count = go(proxy_count)
                str.set('上账单状态更新为‘2’成功，但未找到中间表！')
                if total_count[0][0] == 1:
                    proxy_upsql = "update bank_proxy_payment_received bpr set bpr.vchr_st='902' where bpr.cust_bill_number='"+v1.get()+"'"
                    results_proxy = go(proxy_upsql)
                    str.set('上账单状态更新为‘2’中间表状态更新为‘902’成功，5分钟后定时器重跑同步明细，请注意查看系统明细！！')
                else:
                    proxy_del = "delete from bank_proxy_payment_received bpr where bpr.vchr_st is null and bpr.cust_bill_number='"+v1.get()+"'"
                    results_proxy = go(proxy_del)
                    proxy_upsql = "update bank_proxy_payment_received bpr set bpr.vchr_st='902' where bpr.cust_bill_number='" + v1.get() + "'"
                    results_proxy = go(proxy_upsql)
                    str.set('已删除多余中间表且上账单状态更新为‘2’中间表状态更新为‘902’成功，5分钟后定时器重跑同步明细，请注意查看系统明细！！')

        if v1.get()[0:2] == 'CW':
            withdrawl_upsql = "update capital_withdrawl cw set cw.status = 2 where cw.withdrawl_id='" + v1.get() + "'"
            results_withdrawl = go(withdrawl_upsql)
            if results_withdrawl == 0:
                str.set('未找到下账单！')
            else:
                # 查询中间表是否存在，或者是否存在多条明细，无中间表则返回提示未生成中间表，有多条中间表则删除vchr_st为空的中间表
                proxy_count_cw = "select count(*) from bank_proxy_payment_received bpr where bpr.cust_bill_number='" + v1.get() + "'"
                total_count_cw = go(proxy_count_cw)
                str.set('下账单状态更新为‘2’成功，但未找到中间表！')
                if total_count_cw[0][0] == 1:
                    proxy_upsql_cw = "update bank_proxy_payment_received bpr set bpr.vchr_st='902' where bpr.cust_bill_number='" + v1.get() + "'"
                    results_proxy = go(proxy_upsql_cw)
                    str.set('下账单状态更新为‘2’中间表状态更新为‘902’成功，5分钟后定时器重跑同步明细，请注意查看系统明细！！')
                else:
                    proxy_del_cw = "delete from bank_proxy_payment_received bpr where bpr.vchr_st is null and bpr.cust_bill_number='" + v1.get() + "'"
                    results_proxy = go(proxy_del_cw)
                    proxy_upsql = "update bank_proxy_payment_received bpr set bpr.vchr_st='902' where bpr.cust_bill_number='" + v1.get() + "'"
                    results_proxy = go(proxy_upsql)
                    str.set('已删除多余中间表且下账单状态更新为‘2’中间表状态更新为‘902’成功，5分钟后定时器重跑同步明细，请注意查看系统明细！！')
        # if v1.get()[0:3] == 'JJK':
        #     jjk_upsql = "update capital_cash cc set cc.status=3 where cc.bill_number in \
        #                 (select ccd.main_single_bill_number from capital_cash_detailed ccd where ccd.status=3)"
        #     results_jjk = go(jjk_upsql)
        if (v1.get()[0:2] != 'CW' and v1.get()[0:2] != 'CD'):
            str.set('单号错误，请输入正确的单号！！')

    Button(top, text='提交处理',command=prisoner_money_status).place(x=310, y=80)




def create_newwin_jjk():
    top = Toplevel()
    top.title('接见款处理')
    top.geometry("700x400")
    ts = tkinter.Label(top, text="处理银行交易成功但是状态为“处理中”且明细未同步的情况！！", font=("隶书", 10), fg="red")
    ts.place(x=1, y=20, anchor='nw')
    tl = tkinter.Label(top, text="以下接见款状态为处理中:", font=("隶书", 10), fg="green")
    tl.place(x=5, y=40, anchor='nw')

    columns_title = ("1", "2", "3","4","5")
    jjk_msg = ttk.Treeview(top, height=10, show="headings", columns=columns_title)  # 表格
    jjk_msg.heading('1', text='主单号')
    jjk_msg.heading('2', text='主单号状态')
    jjk_msg.heading('3', text='分单号', )
    jjk_msg.heading('4', text='分单号状态', )
    jjk_msg.heading('5', text='中间表状态', )
    jjk_msg.column('1', anchor='center')
    jjk_msg.column('2', width=90,anchor='center')
    jjk_msg.column('3', anchor='center')
    jjk_msg.column('4', width=90,anchor='center')
    jjk_msg.column('5', width=90, anchor='center')
    jjk_msg.place(x=10, y=80, anchor='nw')

    jjk_proxy = "select cc.bill_number as 主单号,cc.status as 主单号状态,\
                (select tt.bill_number from capital_cash_detailed tt where tt.main_single_bill_number = cc.bill_number) as 分单号,\
                (select tt.status from capital_cash_detailed tt where tt.main_single_bill_number = cc.bill_number) as 分单号状态, \
                bpr.vchr_st as 中间表状态 from capital_cash cc \
                left join bank_proxy_payment_received bpr on cc.bill_number = bpr.cust_bill_number \
                where cc.bill_number in (select ccd.main_single_bill_number from capital_cash_detailed ccd where ccd.status = 3)"
    jjk_query = go(jjk_proxy)

    x = jjk_msg.get_children()
    for item in x:
        jjk_msg.delete(item)

    if len(jjk_query) > 0:
        j = 0
        for i in jjk_query:
            jjk_msg.insert('', 'end', values=jjk_query[j])
            j=j+1
    else:
        jjk_msg.insert('', 'end', values='暂无结果')

    def prisoner_jjk():
        str = StringVar()
        alert_msg = tkinter.Label(top, textvariable=str, font=("隶书", 10), fg="green")
        alert_msg.place(x=10, y=370, anchor='nw')

        jjk_upsql = "update capital_cash cc set cc.status=3 where cc.bill_number in \
                    (select ccd.main_single_bill_number from capital_cash_detailed ccd where ccd.status=3)"
        jjk_results = go(jjk_upsql)
        if jjk_results == 0:
            str.set('未更新任何数据！！')
        else:
            str.set('已更新主单状态9-->3 请等待5分钟后查看系统明细！！')

    Button(top, text='提交处理', command=prisoner_jjk).place(x=10, y=320)

#银行明细重复检测，创建一个窗口来处理
def create_newwin_bankdet():
    top = Toplevel()
    top.title('银行流水重复')
    top.geometry("830x400")
    ts = tkinter.Label(top, text="检测银行流水重复执行删除！！请谨慎使用！！！", font=("隶书", 10), fg="red")
    ts.place(x=1, y=60, anchor='nw')
    tl = tkinter.Label(top, text="以下银行明细表中重复的数据，请核实:", font=("隶书", 10), fg="green")
    tl.place(x=5, y=90, anchor='nw')

    chf_title = tkinter.Label(top, text="重复明细:", font=("隶书", 10), fg="green")
    chf_title.place(x=300, y=90, anchor='nw')

    str_count = StringVar()
    total_count = tkinter.Label(top, textvariable=str_count, font=("隶书", 10), fg="green")
    total_count.place(x=370, y=90, anchor='nw')

    columns_title_c1 = ("1", "2", "3", "4","5")
    show_msgc1 = ttk.Treeview(top, height=10, show="headings", columns=columns_title_c1)  # 表格

    columns_title_c2 = ("1", "2", "3", "4","5","6")
    show_msgc2 = ttk.Treeview(top, height=10, show="headings", columns=columns_title_c2)  # 表格

    str = StringVar()
    alert_msgc = tkinter.Label(top, textvariable=str, font=("隶书", 10), fg="green")
    alert_msgc.place(x=200, y=350, anchor='nw')

    btn = Button(top, text='删除以上明细', fg="green")


    def selection1():
        if CheckVar1.get() == 1:
            btn['bg'] = 'white'
            str.set("")
            CheckVar2.set(0)
            show_msgc2.place_forget()
            show_msgc1.pack
            show_msgc1.heading('1', text='交易日期')
            show_msgc1.heading('2', text='主机交易号')
            show_msgc1.heading('3', text='银行账号')
            show_msgc1.heading('4', text='资金类型', )
            show_msgc1.heading('5', text='单号', )
            show_msgc1.column('1', width=80, anchor='center')
            show_msgc1.column('4', width=80, anchor='center')
            show_msgc1.place(x=5, y=110, anchor='nw')

            show_msgc1_sql="select t.transaction_date,t.accountant_bill_num,t.cust_acc_num,(select bbt.type_name from base_business_type bbt \
                            where t.sub_type_id = bbt.id) as money_type,t.bank_remark from bank_detailed_info t \
                            where (t.transaction_date, t.accountant_bill_num,t.cust_acc_num, t.bank_remark) in ( \
                            select bdi.transaction_date,bdi.accountant_bill_num,bdi.cust_acc_num,bdi.bank_remark from bank_detailed_info bdi \
                            where  bdi.accountant_bill_num is not null and bdi.acc_detailed_type = 1 \
                            group by bdi.transaction_date,bdi.accountant_bill_num,bdi.cash_manage_bill_num,bdi.cust_acc_num,bdi.bank_remark having count(*)>1) \
                            and  id not in (select min(id) from bank_detailed_info tt where tt.accountant_bill_num is not null and tt.acc_detailed_type = 1 \
                            group by tt.transaction_date,tt.accountant_bill_num,tt.cust_acc_num,tt.bank_remark having count(*)>1)"

            show_msgc1_sql_results = go(show_msgc1_sql)
            str_count.set(len(show_msgc1_sql_results))
            x = show_msgc1.get_children()
            for item in x:
                show_msgc1.delete(item)

            if len(show_msgc1_sql_results) > 0:
                j = 0
                for i in show_msgc1_sql_results:
                    show_msgc1.insert('', 'end', values=show_msgc1_sql_results[j])
                    j = j + 1
            else:
                show_msgc1.insert('', 'end', values='暂无结果')

            def show_msgc1_del(event):
               btn['bg'] = 'orange'
               show_msgc1_sql_results = go(show_msgc1_sql)
               if len(show_msgc1_sql_results) > 0:
                #先创建一个表存数据，在进行删除
                    creat_table_sql = "create table data_temp  as select t.id ,t.transaction_date,t.accountant_bill_num,t.cust_acc_num,(select bbt.type_name from base_business_type bbt \
                                where t.sub_type_id = bbt.id) as money_type,t.bank_remark from bank_detailed_info t \
                                where (t.transaction_date, t.accountant_bill_num,t.cust_acc_num, t.bank_remark) in ( \
                                select bdi.transaction_date,bdi.accountant_bill_num,bdi.cust_acc_num,bdi.bank_remark from bank_detailed_info bdi \
                                where  bdi.accountant_bill_num is not null and bdi.acc_detailed_type = 1 \
                                group by bdi.transaction_date,bdi.accountant_bill_num,bdi.cash_manage_bill_num,bdi.cust_acc_num,bdi.bank_remark having count(*)>1) \
                                and  id not in (select min(id) from bank_detailed_info tt where tt.accountant_bill_num is not null and tt.acc_detailed_type = 1 \
                                group by tt.transaction_date,tt.accountant_bill_num,tt.cust_acc_num,tt.bank_remark having count(*)>1)"

                    creat_table_results = go(creat_table_sql)

                    show_msgc1_del_sql = "delete from bank_detailed_info bdi where bdi.id in (select id from data_temp)"
                    show_msgc1_del_results = go(show_msgc1_del_sql)

                    #删除 data_temp
                    drop_table_sql = "drop table data_temp"
                    drop_results = go(drop_table_sql)
                    str.set("数据删除成功！！！")
                    #删除后再次进行查询是否存在重复明细
                    agin_queryc1_results = go(show_msgc1_sql)

                    x = show_msgc1.get_children()
                    for item in x:
                        show_msgc1.delete(item)

                    if len(agin_queryc1_results) > 0:
                        j = 0
                        for i in agin_queryc1_results:
                            show_msgc1.insert('', 'end', values=agin_queryc1_results[j])
                            j = j + 1
                    else:
                        show_msgc1.insert('', 'end', values='暂无结果')
               else:
                    str.set("没有重复数据，无需操作！！！")

        # btn1 = Button(top, text='删除以上明细',fg="green")
        btn.bind('<Button-1>', show_msgc1_del)
        btn.place(x=5, y=350)

    def selection2():
        if CheckVar2.get() == 1:
            btn['bg'] = 'white'
            str.set("")
            CheckVar1.set(0)
            show_msgc1.place_forget()
            show_msgc2.pack
            show_msgc2.heading('1', text='交易日期')
            show_msgc2.heading('2', text='主机交易号')
            show_msgc2.heading('3', text='合约交易号')
            show_msgc2.heading('4', text='银行账号', )
            show_msgc2.heading('5', text='资金类型', )
            show_msgc2.heading('6', text='银行备注', )
            show_msgc2.column('1', width=80, anchor='center')
            show_msgc2.column('5', width=80, anchor='center')
            show_msgc2.column('6', width=80, anchor='center')
            show_msgc2.place(x=5, y=110, anchor='nw')

            show_msgc2_sql = "select t.transaction_date,t.accountant_bill_num,t.cash_manage_bill_num,t.cust_acc_num,\
                                    (select bbt.type_name from base_business_type bbt where t.sub_type_id = bbt.id) as money_type,t.bank_remark from bank_detailed_info t \
                                    where (t.transaction_date, t.accountant_bill_num,t.cash_manage_bill_num, t.cust_acc_num, t.bank_remark) in ( \
                                    select bdi.transaction_date,bdi.accountant_bill_num,bdi.cash_manage_bill_num,bdi.cust_acc_num,bdi.bank_remark from bank_detailed_info bdi \
                                    where  bdi.cash_manage_bill_num is not null and bdi.acc_detailed_type = 1 group by bdi.transaction_date,bdi.accountant_bill_num,bdi.cash_manage_bill_num,bdi.cust_acc_num,bdi.bank_remark \
                                    having count(*)>1)and  id not in (select min(id) from bank_detailed_info tt \
                                    where tt.cash_manage_bill_num is not null and tt.acc_detailed_type = 1 group by tt.transaction_date,tt.accountant_bill_num,tt.cash_manage_bill_num,tt.cust_acc_num,tt.bank_remark having count(*)>1)"

            show_msgc2_sql_results = go(show_msgc2_sql)
            str_count.set(len(show_msgc2_sql_results))
            x = show_msgc2.get_children()
            for item in x:
                show_msgc2.delete(item)

            if len(show_msgc2_sql_results) > 0:
                j = 0
                for i in show_msgc2_sql_results:
                    show_msgc2.insert('', 'end', values=show_msgc2_sql_results[j])
                    j = j + 1
            else:
                show_msgc1.insert('', 'end', values='暂无结果')

            def show_msgc2_del(event):
                btn['bg'] = 'grey'
                show_msgc2_sql_results = go(show_msgc2_sql)
                if len(show_msgc2_sql_results) > 0:
                    creat_table_sql = "create table data_temp as select t.id,t.transaction_date,t.accountant_bill_num,t.cash_manage_bill_num,t.cust_acc_num,\
                                        (select bbt.type_name from base_business_type bbt where t.sub_type_id = bbt.id) as money_type,t.bank_remark from bank_detailed_info t \
                                        where (t.transaction_date, t.accountant_bill_num,t.cash_manage_bill_num, t.cust_acc_num, t.bank_remark) in ( \
                                        select bdi.transaction_date,bdi.accountant_bill_num,bdi.cash_manage_bill_num,bdi.cust_acc_num,bdi.bank_remark from bank_detailed_info bdi \
                                        where  bdi.cash_manage_bill_num is not null and bdi.acc_detailed_type = 1 group by bdi.transaction_date,bdi.accountant_bill_num,bdi.cash_manage_bill_num,bdi.cust_acc_num,bdi.bank_remark \
                                        having count(*)>1)and  id not in (select min(id) from bank_detailed_info tt \
                                        where tt.cash_manage_bill_num is not null and tt.acc_detailed_type = 1 group by tt.transaction_date,tt.accountant_bill_num,tt.cash_manage_bill_num,tt.cust_acc_num,tt.bank_remark having count(*)>1)"

                    creat_table_results = go(creat_table_sql)

                    show_msgc2_del_sql = "delete from bank_detailed_info bdi where bdi.id in (select id from data_temp)"
                    show_msgc2_del_results = go(show_msgc2_del_sql)

                    # 删除 data_temp
                    drop_table_sql = "drop table data_temp"
                    drop_results = go(drop_table_sql)
                    str.set("数据删除成功！！！")
                    # 删除后再次进行查询是否存在重复明细
                    agin_queryc2_results = go(show_msgc2_sql)

                    x = show_msgc2.get_children()
                    for item in x:
                        show_msgc2.delete(item)

                    if len(agin_queryc2_results) > 0:
                        j = 0
                        for i in agin_queryc2_results:
                            show_msgc2.insert('', 'end', values=agin_queryc2_results[j])
                            j = j + 1
                    else:
                        show_msgc2.insert('', 'end', values='暂无结果')
                else:
                    str.set("没有重复数据，无需操作！！！")

        # btn2 = Button(top, text='删除以上明细', fg="green")
        btn.bind('<Button-1>', show_msgc2_del)
        btn.place(x=5, y=350)


    CheckVar1 = IntVar()
    CheckVar2 = IntVar()
    C1 = Checkbutton(top, text="一户通版本", variable=CheckVar1, command=selection1, \
                     onvalue=1, offvalue=0, height=3, \
                     width=20)
    C2 = Checkbutton(top, text="智慧政法1.0", variable=CheckVar2, command=selection2, \
                     onvalue=1, offvalue=0, height=3, \
                     width=20)
    # C1.pack()
    # C2.pack()

    C1.place(x=150,y=1, anchor='nw')
    C2.place(x=300, y=1, anchor='nw')



def change_status_wid():
    top = Toplevel()
    top.title('上下账单状态修改')
    top.geometry("600x300")
    ts = tkinter.Label(top, text="处理银行未交易成功但是状态为“处理中”且无中间表的情况！！", font=("隶书", 10), fg="red")
    ts.place(x=1, y=20, anchor='nw')
    tl = tkinter.Label(top, text="输入单号:", font=("隶书", 10), fg="green")
    tl.place(x=5, y=60, anchor='nw')

    order = StringVar()
    ordernum = Entry(top, textvariable=order, width=40)
    # bill_number = v1.get()
    ordernum.grid(row=1, column=0, padx=10, pady=10)
    ordernum.place(x=10, y=80)

    def prisoner_money_changestatus():
        str = StringVar()
        alert_msg = tkinter.Label(top, textvariable=str, font=("隶书", 10), fg="green")
        alert_msg.place(x=1, y=150, anchor='nw')
        if order.get()[0:2] == 'CD':
            #先查询输入的单号是否存在
            query_sql = "select * from capital_deposit cd where cd.deposit_id='" + order.get() + "'"
            query_sql_results = go(query_sql)
            if len(query_sql_results) > 0:
                #更新主单
                cd_sql = "update capital_deposit cd set cd.status = 5 where cd.status in (9,2) and cd.deposit_id not in \
                            (select bpr.cust_bill_number from bank_proxy_payment_received bpr ) and cd.deposit_id='" + order.get() + "'"
                cd_sql_results = go(cd_sql)
                #更新明细单
                cd_d_sql = "update capital_deposit_item cdi set cdi.status = 0 where cdi.status = 2 and \
                                cdi.deposit_id not in (select bpr.cust_bill_number from bank_proxy_payment_received bpr)and cdi.deposit_id='" + order.get() + "'"
                cd_d_sql_results = go(cd_d_sql)
                str.set("上账主单和明细单更新成功！！")
            else:
                str.set("未查询到此上账单号！！")
        if order.get()[0:2] == 'CW':
            query_sql = "select * from capital_withdrawl cw where cw.withdrawl_id='" + order.get() + "'"
            query_sql_results = go(query_sql)
            if len(query_sql_results) > 0:
                cw_sql = "update capital_withdrawl cw set cw.status = 5 where cw.status in (9,2) and cw.deposit_id not in \
                            (select bpr.cust_bill_number from bank_proxy_payment_received bpr ) and cw.withdrawl_id='" + order.get() + "'"
                cw_sql_results = go(cw_sql)
                cw_d_sql = "update capital_withdrawl_item cwi set cwi.status = 0 where cwi.status = 2 and \
                            cwi.withdrawl_id not in (select bpr.cust_bill_number from bank_proxy_payment_received bpr)and cwi.withdrawl_id='" + order.get() + "'"
                cw_d_sql_results = go(cw_d_sql)
                str.set("下账主单和明细单更新成功！！")
            else:
                str.set("未查询到此下账单号！！")
        if order.get()[0:3] == 'JJK':
            query_sql = "select * from capital_cash_detailed ccd where ccd.bill_number='" + order.get() + "'"
            query_sql_results = go(query_sql)
            if len(query_sql_results) > 0:
                ccd_sql = "update capital_cash_detailed ccd set ccd.status = 0 where ccd.status=3 and ccd.main_single_bill_number \
                                not in (select bpr.cust_bill_number from bank_proxy_payment_received bpr)and ccd.bill_number = '" + order.get() + "'"
                ssd_sql_results = go(ccd_sql)
                str.set("该接见款状态修改成功！！！")
            else:
                str.set("未查询到此接见款单号！！")

        if order.get()[0:2] not in  ('CD','CD') and order.get()[0:3] != 'JJK':
            str.set("单号错误！！")

    Button(top, text='提交处理', command=prisoner_money_changestatus).place(x=310, y=80)

def change_prisoner_bh():
    top = Toplevel()
    top.title('罪犯编号修改')
    top.geometry("600x300")
    ts = tkinter.Label(top, text="将修改所涉及的表：base_offender_info，offender_encounterhistory，", font=("隶书", 10), fg="red")
    ts.place(x=1, y=20, anchor='nw')
    t1 = tkinter.Label(top, text="输入当前编号:", font=("隶书", 10), fg="green")
    t1.place(x=5, y=60, anchor='nw')

    t2 = tkinter.Label(top, text="输入更改的编号:", font=("隶书", 10), fg="green")
    t2.place(x=5, y=105, anchor='nw')

    curr_bh = StringVar()
    curr = Entry(top, textvariable=curr_bh, width=40)
    # bill_number = v1.get()
    curr.grid(row=1, column=0, padx=10, pady=10)
    curr.place(x=10, y=80)

    change_bh = StringVar()
    changebh = Entry(top, textvariable=change_bh, width=40)
    # bill_number = v1.get()
    changebh.grid(row=1, column=0, padx=10, pady=10)
    changebh.place(x=10, y=125)

    msg_show_str1 = StringVar()
    msg_show1 = tkinter.Label(top, textvariable=msg_show_str1, font=("隶书", 10), fg="green")
    msg_show1.place(x=10, y=200, anchor='nw')

    msg_show_str2 = StringVar()
    msg_show2 = tkinter.Label(top, textvariable=msg_show_str2, font=("隶书", 10), fg="green")
    msg_show2.place(x=10, y=230, anchor='nw')

    def prisoner_bh_change():
        #首先查询要更改的编号是否已经在使用
        query_bh_sql = "select boi.zf_bh from base_offender_info boi where boi.zf_bh='"+change_bh.get() + "'"
        query_bh_sql_results = go(query_bh_sql)
        if len(query_bh_sql_results) != 0:
            msg_show_str1.set(change_bh.get())
            msg_show_str2.set("编号已经存在，禁止更改！！！")
        else:
            boi_update_sql = "update base_offender_info boi set boi.zf_bh='"+change_bh.get() + "'"+"where boi.zf_bh = '" +curr_bh.get() + "'"
            boi_update_sql_results = go(boi_update_sql)
            if boi_update_sql_results != 0:
                msg_show_str1.set("base_offender_info更改成功......")
            else:
                msg_show_str1.set("base_offender_info未找到数据，更改失败......")

            oe_update_sql = "update offender_encounterhistory oe set oe.offender_code='"+change_bh.get() + "'"+" where oe.offender_code = '" +curr_bh.get() + "'"
            oe_update_sql_results = go(oe_update_sql)
            if oe_update_sql_results != 0:
                msg_show_str2.set("offender_encounterhistory更改成功......")
            else:
                msg_show_str2.set("offender_encounterhistory未找到数据，更改失败......")

    Button(top, text='提交更改', command=prisoner_bh_change).place(x=10, y=160)






#一下一级窗口
comvalue=tkinter.StringVar()#窗体自带的文本，新建一个值
comboxlist=ttk.Combobox(win,width=65, height=20,font=1,textvariable=comvalue) #初始化
comboxlist["values"]=msg
comboxlist.current(0) #默认选择第一个
# comboxlist.bind("<<ComboboxSelected>>",go) #绑定事件,(下拉列表框被选中时，绑定go()函数)
comboxlist.place(x=10,y=40,anchor='nw')

# link_btn=tkinter.Button(win,text="确定", font=("隶书",10),command=go,width=6, height=1, fg="green")
# link_btn.place(x=400, y=40, anchor='nw')
# comboxlist.pack()

#功能----上下账模块，接见款
w2 = tkinter.Label(win, text="1、检测：状态为0,2,9的；2、上下账处理：处理银行交易成功且未同步的明细的；3、状态修改：处理银行未交易且没有中间表的单子为待提交（含接见款）",font=("隶书",9), fg="red")
w2.place(x=10, y=130, anchor='nw')


button1 = tkinter.Button(win, text="上下账单检测", font=("隶书",10), command=prisoner_money_check,width=12, height=2, fg="green")
button1.place(x=10, y=90, anchor='nw')
# button1.pack()

button2 = tkinter.Button(win, text="上下账单处理", font=("隶书",10), command=create_newwin,width=12, height=2, fg="green")
button2.place(x=115, y=90, anchor='nw')

button3 = tkinter.Button(win, text="上下账处理中状态修改", font=("隶书",10), command=change_status_wid,width=18, height=2, fg="green")
button3.place(x=220, y=90, anchor='nw')

button4 = tkinter.Button(win, text="接见款处理中", font=("隶书",10), command=create_newwin_jjk, width=12, height=2, fg="green")
button4.place(x=370, y=90, anchor='nw')

button5 = tkinter.Button(win, text="罪犯编号修改", font=("隶书",10), command=change_prisoner_bh,width=12,height=2, fg="green")
button5.place(x=480, y=90, anchor='nw')

b1 = tkinter.Label(win, text="返回信息：",font=("隶书",10), fg="green")
b1.place(x=5, y=200, anchor='nw')

# return_msg=tkinter.Text(win,width=100,height=10)
columns = ("1", "2","3")
return_msg=ttk.Treeview(win, height=6, show="headings", columns=columns)  # 表格
return_msg.heading('1',text='单号')
return_msg.heading('2',text='银行账号')
return_msg.heading('3',text='状态',)
return_msg.column('1',anchor='center')
return_msg.column('2',anchor='center')
return_msg.column('3',width=50,anchor='center')
return_msg.place(x=80, y=150, anchor='nw')

bil = tkinter.Label(win, text="单号：",font=("隶书",10), fg="green")
bil.place(x=550, y=200, anchor='nw')
bill_num = tkinter.Text(win,width=30,height=11)
bill_num.config(wrap=WORD)
bill_num.place(x=600, y=150, anchor='nw')



#功能---服务器模块
button6 = tkinter.Button(win, text="服务器空间查询", command=link_linux_server,font=("隶书",10), width=13, height=2, fg="green")
button6.place(x=10, y=300, anchor='nw')

button7 = tkinter.Button(win, text="tomcat日志清除", font=("隶书",10), width=14, height=2, fg="green")
button7.place(x=130, y=300, anchor='nw')

#数据重复
button8 = tkinter.Button(win, text="银行流水重复删除", command=create_newwin_bankdet,font=("隶书",10), width=15, height=2, fg="green")
button8.place(x=260, y=300, anchor='nw')

b2 = tkinter.Label(win, text="返回信息：",font=("隶书",10), fg="green")
b2.place(x=20, y=400, anchor='nw')


button2 = tkinter.Button(win, text="导出生成Excel",font=("隶书",15), width=15, height=3, fg="green")
button2.place(x=300, y=600, anchor='nw')

#连接oracle



win.mainloop()