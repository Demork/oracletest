import tkinter
from tkinter import ttk
import cx_Oracle as oea
from tkinter import *
import xlwt
from datetime import datetime
from ssh import SSH
import tkinter.filedialog
import unicodedata
from tkinter import messagebox  # 弹出提示框

db_port = '1521'
ser_name = 'pora12c1.lecent.domain'
msg=[]
bill_number=''
server_msg=[]
class link_oracle:
    f = open("C:\\lecent_all\\oracle_link.txt","r",encoding='UTF-8')
    lines = f.readlines()
    for line in lines:
        line = line.rstrip("\n")
        msg.append(line)
class link_server:
    f = open("C:\\lecent_all\\linux_server.txt","r",encoding='UTF-8')
    lines = f.readlines()
    for line in lines:
        line = line.rstrip("\n")
        server_msg.append(line)

win = tkinter.Tk()
win.title("老资管常见数据问题处理脚本工具(oracle)")
win.geometry("900x500")
win.resizable(0, 0)  # 锁定窗口
#说明
w1 = tkinter.Label(win, text="以下功能直接点击后会自动进行检查，未返回结果之前请勿关闭！！",font=("隶书",9), fg="red")
w1.place(x=250, y=5, anchor='nw')

#连接oracle
def go(sql):  #处理事件，*args表示可变参数
    # print(comboxlist.get().split(',')) #打印选中的值
    #获取IP,用户及密码连接Oracle
    IP=comboxlist.get().split(',')[0]
    user=comboxlist.get().split(',')[1]
    psd=comboxlist.get().split(',')[2]
    print('ip:'+IP+'用户:'+user+'密码:'+psd)
    conn = oea.connect(user,psd,IP + ':' + '46108' + '/' + 'pora12c1.lecent.domain')
    # conn = oea.connect(user, psd, IP + ':' + '1521' + '/' + 'pora12c1.lecent.domain')
    print(conn)
    aletr_msg = tkinter.Label(win, text="数据连接成功...", font=("隶书", 11), fg="green")
    aletr_msg.place(x=650, y=40, anchor='nw')
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

#查询上下账状态不正常的
def prisoner_money_check():
    sql = "select * from (select t.deposit_id as 单号,t.bank_card_number as 卡号,t.status as 状态 from capital_deposit_item t \
    where t.status in (0,2,9) union all select cwi.withdrawl_id as 单号,cwi.bank_card_number as 卡号,cwi.status as 状态 \
    from capital_withdrawl_item cwi where cwi.status in (0,2,9)  union all select odi.bill_number as 单号,odi.bank_card_number as 卡号,111 as 状态 \
    from offender_detailed_info odi where (odi.bill_number like 'CW%' or odi.bill_number like 'CD%') and odi.offender_trasaction_money <>0 \
    and odi.bill_number not in (select bill from (select cdi.deposit_id as bill  from  capital_deposit cdi union all \
    select t.withdrawl_id as bill from capital_withdrawl t) )) order by 状态 desc "
    results = go(sql)

    bill_num.delete("1.0", tkinter.END)
    sql_bill = "select 单号 from (select t.deposit_id as 单号,t.bank_card_number as 卡号,t.status as 状态 from capital_deposit_item t \
    where t.status in (0,2,9) union all select cwi.withdrawl_id as 单号,cwi.bank_card_number as 卡号,cwi.status as 状态 \
    from capital_withdrawl_item cwi where cwi.status in (0,2,9)union all select odi.bill_number as 单号,odi.bank_card_number as 卡号,111 as 状态 \
    from offender_detailed_info odi where (odi.bill_number like 'CW%' or odi.bill_number like 'CD%') and odi.offender_trasaction_money <>0 \
    and odi.bill_number not in (select bill from (select cdi.deposit_id as bill  from  capital_deposit cdi union all \
    select t.withdrawl_id as bill from capital_withdrawl t) )) group by 单号  "
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

#上下账处理
def create_newwin():
    top = Toplevel()
    top.title('上下账处理')
    top.geometry("600x300")
    top.resizable(0, 0)  # 锁定窗口
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

#接见款处理
def create_newwin_jjk():
    top = Toplevel()
    top.title('接见款处理')
    top.geometry("700x400")
    top.resizable(0, 0)  # 锁定窗口
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
    top.geometry("850x400")
    top.resizable(0, 0)  # 锁定窗口
    ts = tkinter.Label(top, text="检测银行流水重复执行删除！！请谨慎使用！！！", font=("隶书", 10), fg="red")
    ts.place(x=1, y=60, anchor='nw')
    tl = tkinter.Label(top, text="以下银行明细表中重复的数据，请核实:", font=("隶书", 10), fg="green")
    tl.place(x=5, y=90, anchor='nw')

    chf_title = tkinter.Label(top, text="重复明细:", font=("隶书", 10), fg="green")
    chf_title.place(x=300, y=90, anchor='nw')

    chf_title = tkinter.Label(top, text="总金额:", font=("隶书", 10), fg="green")
    chf_title.place(x=450, y=90, anchor='nw')

    money_count = StringVar()
    money_total = tkinter.Label(top, textvariable=money_count, font=("隶书", 10), fg="green")
    money_total.place(x=500, y=90, anchor='nw')

    str_count = StringVar()
    total_count = tkinter.Label(top, textvariable=str_count, font=("隶书", 10), fg="green")
    total_count.place(x=370, y=90, anchor='nw')

    columns_title_c1 = ("1", "2", "3", "4","5")
    show_msgc1 = ttk.Treeview(top, height=10, show="headings", columns=columns_title_c1)  # 表格

    columns_title_c2 = ("1", "2", "3", "4","5","6")
    show_msgc2 = ttk.Treeview(top, height=10, show="headings", columns=columns_title_c2)  # 表格

    str = StringVar()
    alert_msgc = tkinter.Label(top, textvariable=str, font=("隶书", 10), fg="green")
    alert_msgc.place(x=400, y=350, anchor='nw')

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
            show_msgc2.heading('2', text='交易金额')
            show_msgc2.heading('3', text='合约交易号')
            show_msgc2.heading('4', text='银行账号', )
            show_msgc2.heading('5', text='资金类型', )
            show_msgc2.heading('6', text='银行备注', )
            show_msgc2.column('1', width=80, anchor='center')
            show_msgc2.column('2', width=60, anchor='center')
            show_msgc2.column('5', width=80, anchor='center')
            show_msgc2.column('6', width=80, anchor='center')
            show_msgc2.place(x=5, y=110, anchor='nw')

            show_msgc2_sql = "select t.transaction_date,t.transaction_money,t.cash_manage_bill_num,t.cust_acc_num,\
                                    (select bbt.type_name from base_business_type bbt where t.sub_type_id = bbt.id) as money_type,t.bank_remark from bank_detailed_info t \
                                    where (t.transaction_date,t.cash_manage_bill_num, t.cust_acc_num, t.bank_remark) in ( \
                                    select bdi.transaction_date,bdi.cash_manage_bill_num,bdi.cust_acc_num,bdi.bank_remark from bank_detailed_info bdi \
                                    where  bdi.cash_manage_bill_num is not null and bdi.acc_detailed_type = 1 group by bdi.transaction_date,bdi.accountant_bill_num,bdi.cash_manage_bill_num,bdi.cust_acc_num,bdi.bank_remark \
                                    having count(*)>1)and  id not in (select min(id) from bank_detailed_info tt \
                                    where tt.cash_manage_bill_num is not null and tt.acc_detailed_type = 1 group by tt.transaction_date,tt.accountant_bill_num,tt.cash_manage_bill_num,tt.cust_acc_num,tt.bank_remark having count(*)>1) order by t.transaction_date asc "

            show_msgc2_sql_results = go(show_msgc2_sql)
            str_count.set(len(show_msgc2_sql_results))
            x = show_msgc2.get_children()
            for item in x:
                show_msgc2.delete(item)
            TotalMoney = 0
            if len(show_msgc2_sql_results) > 0:
                j = 0
                for i in show_msgc2_sql_results:
                    show_msgc2.insert('', 'end', values=show_msgc2_sql_results[j])
                    TotalMoney = TotalMoney + show_msgc2_sql_results[j][1]
                    j = j + 1
            else:
                show_msgc2.insert('', 'end', values='暂无结果')
            money_count.set(round(TotalMoney,2))

            def show_msgc2_del(event):
                btn['bg'] = 'grey'
                show_msgc2_sql_results = go(show_msgc2_sql)
                if len(show_msgc2_sql_results) > 0:
                    creat_table_sql = "create table data_temp as select t.id,t.transaction_date,t.transaction_money,t.cash_manage_bill_num,t.cust_acc_num,\
                                        (select bbt.type_name from base_business_type bbt where t.sub_type_id = bbt.id) as money_type,t.bank_remark from bank_detailed_info t \
                                        where (t.transaction_date,t.cash_manage_bill_num, t.cust_acc_num, t.bank_remark) in ( \
                                        select bdi.transaction_date,bdi.cash_manage_bill_num,bdi.cust_acc_num,bdi.bank_remark from bank_detailed_info bdi \
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
                    str_count.set(len(agin_queryc2_results))
                    x = show_msgc2.get_children()
                    for item in x:
                        show_msgc2.delete(item)
                    AginSumMoney = 0
                    if len(agin_queryc2_results) > 0:
                        j = 0
                        for i in agin_queryc2_results:
                            show_msgc2.insert('', 'end', values=agin_queryc2_results[j])
                            AginSumMoney = AginSumMoney + show_msgc2_sql_results[j][1]
                            j = j + 1
                    else:
                        show_msgc2.insert('', 'end', values='暂无结果')
                else:
                    str.set("没有重复数据，无需操作！！！")
                money_count.set(round(AginSumMoney, 2))

            # 导出btn
            def report_data():
                file = xlwt.Workbook()
                sheet1 = file.add_sheet('sheet1', cell_overwrite_ok=True)
                list_data = []
                for i in range(0, len(columns_title_c2)):
                    sheet1.write(0, i, show_msgc2.heading(i)['text'])
                for j in show_msgc2_sql_results:
                    list_data.append(j)
                b = 1
                for k, q in enumerate(list_data):
                    for a in range(0, len(list_data[k])):
                        sheet1.write(b, a, list_data[k][a])
                    b = b + 1
                datestring = datetime.strftime(datetime.now(), ' %Y-%m-%d %H-%M-%S')
                filename = tkinter.filedialog.asksaveasfilename(filetypes=[('xlsx', '*.xlsx')],
                                                                initialfile="重复明细检测" + datestring, initialdir='C:\\')
                filename = filename + '.xls'
                file.save(filename)

            Button(top, text='导出数据', command=report_data, fg='green').place(x=150, y=350)

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

    C1.place(x=150,y=1, anchor='nw')
    C2.place(x=300, y=1, anchor='nw')

#上下账单状态修改，未交易
def change_status_wid():
    top = Toplevel()
    top.title('上下账单状态修改')
    top.geometry("600x300")
    top.resizable(0, 0)  # 锁定窗口
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
                cd_sql = "update capital_deposit cd set cd.status = 5 where cd.status in (9,2,3) and cd.deposit_id not in \
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
                cw_sql = "update capital_withdrawl cw set cw.status = 5 where cw.status in (9,2,3) and cw.deposit_id not in \
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

#罪犯编号修改
def change_prisoner_bh():
    top = Toplevel()
    top.title('罪犯编号修改')
    top.geometry("600x300")
    top.resizable(0, 0)  # 锁定窗口
    ts = tkinter.Label(top, text="将修改所涉及的表：base_offender_info，offender_encounterhistory", font=("隶书", 10), fg="red")
    ts.place(x=1, y=5, anchor='nw')
    info = tkinter.Label(top, text="编号更改:未使用编号修改。编号互换:两罪犯编号之间进行互换", font=("隶书", 10), fg="red")
    info.place(x=1, y=25, anchor='nw')
    t1 = tkinter.Label(top, text="输入当前编号:", font=("隶书", 10), fg="green")
    t1.place(x=5, y=60, anchor='nw')

    t2 = tkinter.Label(top, text="输入更改的编号:", font=("隶书", 10), fg="green")
    t2.place(x=5, y=105, anchor='nw')

    curr_bh = StringVar()
    curr = Entry(top, textvariable=curr_bh, width=30)
    curr.grid(row=1, column=0, padx=10, pady=10)
    curr.place(x=10, y=80)

    change_bh = StringVar()
    changebh = Entry(top, textvariable=change_bh, width=30)
    changebh.grid(row=1, column=0, padx=10, pady=10)
    changebh.place(x=10, y=125)

    msg_show_str1 = StringVar()
    msg_show1 = tkinter.Label(top, textvariable=msg_show_str1, font=("隶书", 10), fg="green")
    msg_show1.place(x=10, y=200, anchor='nw')

    msg_show_str2 = StringVar()
    msg_show2 = tkinter.Label(top, textvariable=msg_show_str2, font=("隶书", 10), fg="green")
    msg_show2.place(x=10, y=230, anchor='nw')

    msg_show_str3 = StringVar()
    msg_show3 = tkinter.Label(top, textvariable=msg_show_str3, font=("隶书", 10), fg="green")
    msg_show3.place(x=250, y=80, anchor='nw')

    msg_show_str4 = StringVar()
    msg_show4 = tkinter.Label(top, textvariable=msg_show_str4, font=("隶书", 10), fg="green")
    msg_show4.place(x=250, y=125, anchor='nw')

    msg_show_str5 = StringVar()
    msg_show5 = tkinter.Label(top, textvariable=msg_show_str5, font=("隶书", 10), fg="green")
    msg_show5.place(x=400, y=80, anchor='nw')

    msg_show_str6 = StringVar()
    msg_show6 = tkinter.Label(top, textvariable=msg_show_str6, font=("隶书", 10), fg="green")
    msg_show6.place(x=400, y=125, anchor='nw')

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

    def exchange_bh():
        query_sql_first = "select boi.xm,boi.zf_bh from base_offender_info boi where boi.zf_bh='"+curr_bh.get() + "'" # 当前罪犯信息
        query_sql_first_msg = go(query_sql_first)
        msg_show_str3.set(query_sql_first_msg)
        query_sql_sed = "select boi.xm,boi.zf_bh from base_offender_info boi where boi.zf_bh='" + change_bh.get() + "'" # 第二个罪犯信息
        query_sql_sed_msg = go(query_sql_sed)
        msg_show_str4.set(query_sql_sed_msg)

        if query_sql_first_msg and query_sql_sed_msg:
            #执行修改程序,第一个
            first_boi_update = "update base_offender_info boi set boi.zf_bh='"+query_sql_sed_msg[0][1]+ "'"+" where boi.zf_bh = '"+query_sql_first_msg[0][1]+ "'"
            f_boi_rel = go(first_boi_update)
            first_oe_update = "update offender_encounterhistory oe set oe.offender_code='"+change_bh.get() + "'"+" where oe.offender_code = '" +query_sql_first_msg[0][1] + "'"
            f_oe_rel = go(first_oe_update)
            # 执行修改程序,第二个
            sed_boi_update = "update base_offender_info boi set boi.zf_bh='"+query_sql_first_msg[0][1] + "'"+" where boi.zf_bh = '" +query_sql_sed_msg[0][1]+ "'"+" and boi.xm="+"'"+query_sql_sed_msg[0][0]+"'"
            s_boi_rel = go(sed_boi_update)
            sed_oe_update = "update offender_encounterhistory oe set oe.offender_code='"+curr_bh.get() + "'"+" where oe.offender_code = '" +query_sql_sed_msg[0][1]+ "'"+" and oe.offender_name="+"'"+query_sql_sed_msg[0][0]+"'"
            s_oe_rel = go(sed_oe_update)
            if f_boi_rel and s_boi_rel:
                msg_show_str1.set("base_offender_info表提交互换成功，请查看系统......")
            if first_oe_update != 0 or sed_oe_update !=0:
                msg_show_str2.set("offender_encounterhistory表更新成功，请查看系统......")

            #第二次查询修改后的结果
            query_update_rel_f = "select boi.xm,boi.zf_bh from base_offender_info boi where boi.zf_bh='"+curr_bh.get()+ "'"
            query_update_rul_f = go(query_update_rel_f)
            msg_show_str5.set(query_update_rul_f)
            query_update_rel_s = "select boi.xm,boi.zf_bh from base_offender_info boi where boi.zf_bh='" +change_bh.get()+ "'"
            query_update_rul_s = go(query_update_rel_s)
            msg_show_str6.set(query_update_rul_s)

    Button(top, text='编号更改', command=prisoner_bh_change).place(x=10, y=160)
    Button(top, text='编号互换', command=exchange_bh).place(x=100, y=160)

#连接服务器，获取配置文件IP信息
def clear_catalina():
    top = Toplevel()
    top.title('项目日志检测')
    top.geometry("800x400")
    # top.resizable(0, 0)  # 锁定窗口
    explain = tkinter.Label(top, text="选择服务器查询，点击要清除的路径确认清除即可。。。", font=("隶书", 10), fg="red")
    explain.place(x=1, y=10, anchor='nw')
    server_info = tkinter.StringVar()  # 窗体自带的文本，新建一个值
    server_infolist = ttk.Combobox(top, width=50, height=20, font=1, textvariable=server_info)  # 初始化
    server_infolist["values"] = server_msg
    server_infolist.current()  # 默认选择第一个
    server_infolist.place(x=10, y=40, anchor='nw')

    linux_space_info = tkinter.Text(top, width=55, height=17)
    linux_space_info.config(wrap=WORD)
    linux_space_info.place(x=400, y=110, anchor='nw')

    log_title = ("0", "1")
    log_msg = ttk.Treeview(top, height=10, show="headings", columns=log_title)  # 表格
    log_msg.place(x=10, y=110)
    log_msg.heading('0', text='size')
    log_msg.heading('1', text='path')
    log_msg.column('0', width=50, anchor='center')
    log_msg.column('1', width=300, anchor='center')
    def clear():
        IP = server_infolist.get().split(',')[0]
        port = server_infolist.get().split(',')[1]
        psd = server_infolist.get().split(',')[3]
        params = {
            'hostname': IP,
            'username': 'root',
            'port': port,
            'password': psd,
        }
        print(params)

        with SSH(**params) as ssh:
            stdin, stdout, stderr = ssh.exec_command("find  / -iname 'catalina.out'")
            path_log = stdout.read().decode('utf-8')
            path_log_list = path_log.replace("\n", " ").split(' ')
            print(path_log_list)
            list_msg = []
            for item in path_log_list:
                if item != '':
                    stdin, stdout, stderr = ssh.exec_command("du -sh "+item)
                    return_data = stdout.read().decode('utf-8')
                    # list_msg.append(return_data.replace("\t", " ").split('\n'))
                    for item_msg in return_data.replace("\t", " ").split('\n'):
                        if item_msg != '':
                            list_msg.append(item_msg)
            print(list_msg)

            x = log_msg.get_children()
            for item in x:
                log_msg.delete(item)
            if len(list_msg) > 0:
                j = 0
                for i in list_msg:
                    log_msg.insert('', 'end', values=list_msg[j])
                    j = j + 1
            else:
                log_msg.insert('', 'end', values='暂无结果')

        def treeviewClick(event):  # 单击
            print('单击')
            for item in log_msg.selection():
                item_text = log_msg.item(item, "values")
                mes = messagebox.askyesno('提示', "要执行此清除日志的操作吗？"+"清除路径：\n"+item_text[1])
                if mes is True:
                    with SSH(**params) as ssh:
                        stdin, stdout, stderr = ssh.exec_command("echo -n "" > "+item_text[1])
                        messagebox.showinfo('提示',"已清空日志\n"+item_text[1])
        log_msg.bind('<ButtonRelease-1>', treeviewClick)  # 绑定单击离开事件===========

    def link_linux_server():
        linux_space_info.delete("1.0", tkinter.END)
        IP = server_infolist.get().split(',')[0]
        port = server_infolist.get().split(',')[1]
        psd = server_infolist.get().split(',')[3]
        params = {
            'hostname': IP,
            'username': 'root',
            'port': port,
            'password': psd,
        }
        print(params)
        with SSH(**params) as ssh:
            stdin, stdout, stderr = ssh.exec_command("df -h")
            out_results = stdout.read()
            return_str = str(out_results, encoding="utf-8")
            linux_space_info.insert('insert', return_str)
            # 获取项目路径以及catalina.out  ps -ef | grep tomcat| grep -v grep|awk '{print $2}',ls -l /proc/82073/cwd
            stdin, stdout, stderr = ssh.exec_command("ps -ef | grep tomcat| grep -v grep|awk '{print $2}'")
            pid = stdout.read()
            pid_str = str(pid, encoding="utf-8")
            list_pid = []
            for i in pid_str.split('\n'):
                if i != '':
                    list_pid.append(i)
            # 获取项目路径 和端口
            datalist = []
            for tm_pid in list_pid:
                stdin, stdout, stderr = ssh.exec_command("ls -l /proc/" + tm_pid + "/cwd | awk '{print $11}' &&  netstat -nap | grep " + tm_pid + " | awk '{print $4}' | sed -n '1p' | sed 's/::://'")
                tm_pid_rul = stdout.read().decode('utf-8')
                data = tm_pid_rul.replace("\n", " ").split(' ')
                if isinstance(data, list):
                    if len(data) > 2:
                        if data[-1] == '':
                            data.pop(-1)
                        datalist.append({"path": data[0], "port": data[1]})
                    else:
                        datalist.append({"path": None, "port": None})
            print(datalist)

    Button(top, text="确定", font=("隶书",10),command=clear,width=10, height=2, fg="green").place(x=10, y=70, anchor='nw')
    Button(top, text="服务器空间查询", command=link_linux_server, font=("隶书", 10), width=13, height=2,fg="green").place(x=400, y=70, anchor='nw')


#检测账号资金
def capital_test():
    top = Toplevel()
    top.title('资金检测')
    top.geometry("700x400")
    top.resizable(0, 0) #锁定窗口
    til = tkinter.Label(top, text="1、银行余额检测:一户通专用，检测账户表银行余额是否与历史余额相符(此工具请先执行系统操作的1,2,3步骤)", font=("隶书", 10), fg="red")
    til_1 = tkinter.Label(top,text="2、智慧政法1.0检测:检测账户表银行余额是否与银行流水明细交易金额累加余额相符", font=("隶书", 10), fg="red")
    til_2 = tkinter.Label(top, text="3、系统操作：1、先拉取银行流水表为正常;2、再同步拉取账户余额;3、再执行此工具;4、时间为当天时间 ", font=("隶书", 10), fg="red")
    til_3 = tkinter.Label(top, text="资余额金检测，罪犯余额检测工具通用", font=("隶书", 10), fg="red")
    til.place(x=1, y=1, anchor='nw')
    til_1.place(x=1, y=20, anchor='nw')
    til_2.place(x=1, y=40, anchor='nw')
    til_3.place(x=1, y=60, anchor='nw')

    capital_title = ("0", "1", "2", "3", "4", "5")
    capital_msg = ttk.Treeview(top, height=10, show="headings", columns=capital_title)  # capital_check表格

    bank_msg= ttk.Treeview(top, height=10, show="headings", columns=capital_title)  # bank_check表格

    balance_msg = ttk.Treeview(top, height=10, show="headings", columns=capital_title)  # balance_check表格
    zfzfh_bank_msg = ttk.Treeview(top, height=10, show="headings", columns=capital_title)  # zfzfh_bank_check表格

    capital_total = tkinter.Label(top, text="数量：", font=("隶书", 10), fg="green")
    capital_total.place(x=10, y=360, anchor='nw')
    capital_count = StringVar()
    counts = tkinter.Label(top, textvariable=capital_count, font=("隶书", 10), fg="green")
    counts.place(x=40, y=360, anchor='nw')

    def capital_check():
        bank_msg.place_forget()
        balance_msg.place_forget()
        zfzfh_bank_msg.pack_forget()
        capital_msg.pack()
        capital_msg.place(x=2, y=120)
        capital_msg.heading('0', text='姓名')
        capital_msg.heading('1', text='账号')
        capital_msg.heading('2', text='狱内银行余额')
        capital_msg.heading('3', text='狱外银行余额')
        capital_msg.heading('4', text='差额(狱内-狱外)')
        capital_msg.heading('5', text='账号状态')
        capital_msg.column('0', width=80, anchor='center')
        capital_msg.column('2', width=80, anchor='center')
        capital_msg.column('3', width=80, anchor='center')
        capital_msg.column('4', width=95, anchor='center')
        capital_msg.column('5', width=80, anchor='center')
        capital_query_sql = "select * from (select (select xm from base_offender_info boi where boi.id = t.offender_id) as xm,\
                                t.bank_card_number,t.bank_balance,c.balance,(t.bank_balance - c.balance) as ce,\
                                (case (select status from base_offender_info boi where boi.id = t.offender_id)when 1 then '在狱' else '出狱'end) as status \
                                from base_offender_account t,(select tt.balance, tt.bank_card_number from bank_card_balance tt where tt.create_date >= to_date(to_char(sysdate,'yyyy-mm-dd') ,'yyyy-mm-dd hh24:mi:ss')) c \
                                where t.bank_card_number = c.bank_card_number and t.bank_balance != c.balance) where ce <> 0 order by ce desc"

        capital_query_sql_results = go(capital_query_sql)

        x = capital_msg.get_children()
        for item in x:
            capital_msg.delete(item)
        if len(capital_query_sql_results) > 0:
            capital_count.set(len(capital_query_sql_results))
            j = 0
            for i in capital_query_sql_results:
                capital_msg.insert('', 'end', values=capital_query_sql_results[j])
                j = j + 1
        else:
            capital_msg.insert('', 'end', values='暂无结果')
            capital_count.set('0')

        # 导出btn
        def report_data():
            file = xlwt.Workbook()
            sheet1 = file.add_sheet('sheet1', cell_overwrite_ok=True)
            list_data = []
            for i in range(0, len(capital_title)):
                sheet1.write(0, i, capital_msg.heading(i)['text'])
            for j in capital_query_sql_results:
                list_data.append(j)
            b = 1
            for k, q in enumerate(list_data):
                for a in range(0, len(list_data[k])):
                    sheet1.write(b, a, list_data[k][a])
                b = b + 1
            datestring = datetime.strftime(datetime.now(), ' %Y-%m-%d %H-%M-%S')
            filename = tkinter.filedialog.asksaveasfilename(filetypes=[('xlsx', '*.xlsx')],
                                                            initialfile="银行余额检测" + datestring, initialdir='C:\\')
            filename = filename + '.xls'
            # messagebox.showinfo("提示", "保存完毕~！！！")
            # file.save('E:\\lecent_all\\'+'罪犯余额检测'+datestring+'.xlsx')
            file.save(filename)

        Button(top, text='导出数据', command=report_data, fg='green').place(x=600, y=355)

    def bank_check():
        capital_msg.place_forget()
        balance_msg.place_forget()
        zfzfh_bank_msg.pack_forget()
        bank_msg.pack()
        bank_msg.place(x=2, y=120)
        bank_msg.heading('0', text='姓名')
        bank_msg.heading('1', text='账号')
        bank_msg.heading('2', text='账户表银行余额')
        bank_msg.heading('3', text='资金表累加余额')
        bank_msg.heading('4', text='差额(账户表-资金表)')
        bank_msg.heading('5', text='账号状态')
        bank_msg.column('0', width=80, anchor='center')
        bank_msg.column('2', width=100, anchor='center')
        bank_msg.column('3', width=100, anchor='center')
        bank_msg.column('4', width=120, anchor='center')
        bank_msg.column('5', width=80, anchor='center')

        bank_msg_sql = "select * from (select (select xm from base_offender_info boi where boi.id=tb.offender_id) as xm ,tb.bank_card_number,tb.bank_balance,tt.je,tb.bank_balance - tt.je as ce, \
                        (case (select status from base_offender_info boi where boi.id=tb.offender_id) when 1 then '在狱'else '出狱' end )as status from base_offender_account tb,\
                        (select t.offender_id, sum(case when t.transaction_code = 0 then -t.transaction_money else t.transaction_money end) as je \
                        from capital_detailed_info t where 1 = 1 group by t.offender_id) tt where tb.offender_id = tt.offender_id and tb.bank_balance <> tt.je) order by ce desc"
        bank_msg_sql_results = go(bank_msg_sql)
        x = bank_msg.get_children()
        for item in x:
            bank_msg.delete(item)
        if len(bank_msg_sql_results) > 0:
            capital_count.set(len(bank_msg_sql_results))
            j = 0
            for i in bank_msg_sql_results:
                bank_msg.insert('', 'end', values=bank_msg_sql_results[j])
                j = j + 1
        else:
            bank_msg.insert('', 'end', values='暂无结果')
            capital_count.set('0')

        # 导出btn
        def report_data():
            file = xlwt.Workbook()
            sheet1 = file.add_sheet('sheet1', cell_overwrite_ok=True)
            list_data = []
            for i in range(0, len(capital_title)):
                sheet1.write(0, i, bank_msg.heading(i)['text'])
            for j in bank_msg_sql_results:
                list_data.append(j)
            b = 1
            for k, q in enumerate(list_data):
                for a in range(0, len(list_data[k])):
                    sheet1.write(b, a, list_data[k][a])
                b = b + 1
            datestring = datetime.strftime(datetime.now(), ' %Y-%m-%d %H-%M-%S')
            filename = tkinter.filedialog.asksaveasfilename(filetypes=[('xlsx', '*.xlsx')],
                                                            initialfile="资金余额检测" + datestring, initialdir='C:\\')
            filename = filename + '.xls'
            # messagebox.showinfo("提示", "保存完毕~！！！")
            # file.save('E:\\lecent_all\\'+'罪犯余额检测'+datestring+'.xlsx')
            file.save(filename)

        Button(top, text='导出数据', command=report_data, fg='green').place(x=600, y=355)

    def balance_check():
        capital_msg.place_forget()
        bank_msg.place_forget()
        zfzfh_bank_msg.pack_forget()
        balance_msg.pack()
        balance_msg.place(x=2, y=120)
        balance_msg.heading('0', text='姓名')
        balance_msg.heading('1', text='账号')
        balance_msg.heading('2', text='罪犯余额')
        balance_msg.heading('3', text='台账累加余额')
        balance_msg.heading('4', text='差额(账户表-台账)')
        balance_msg.heading('5', text='账号状态')
        balance_msg.column('0', width=80, anchor='center')
        balance_msg.column('2', width=70, anchor='center')
        balance_msg.column('3', width=90, anchor='center')
        balance_msg.column('4', width=115, anchor='center')
        balance_msg.column('5', width=80, anchor='center')

        balance_msg_sql = "select * from (select (select xm from base_offender_info boi where boi.id=tb.offender_id) as xm ,tb.bank_card_number,tb.total_quota,tt.je,tb.total_quota - tt.je as ce,\
                            (case (select status from base_offender_info boi where boi.id=tb.offender_id) when 1 then '在狱'else '出狱' end )as status from base_offender_account tb,\
                            (select t.offender_id, sum(t.offender_trasaction_money) as je from offender_detailed_info t where 1 = 1 and t.offender_trasaction_money <> 0 group by t.offender_id) tt \
                            where tb.offender_id = tt.offender_id and tb.total_quota <> tt.je) order by ce desc"
        balance_msg_sql_results = go(balance_msg_sql)
        x = balance_msg.get_children()
        for item in x:
            balance_msg.delete(item)
        if len(balance_msg_sql_results) > 0:
            capital_count.set(len(balance_msg_sql_results))
            j = 0
            for i in balance_msg_sql_results:
                balance_msg.insert('', 'end', values=balance_msg_sql_results[j])
                j = j + 1
        else:
            balance_msg.insert('', 'end', values='暂无结果')
            capital_count.set('0')
        #导出btn
        def report_data():
           file = xlwt.Workbook()
           sheet1 = file.add_sheet('sheet1',cell_overwrite_ok=True)
           list_data = []
           for i in range(0,len(capital_title)):
               sheet1.write(0,i,balance_msg.heading(i)['text'])
           for j in balance_msg_sql_results:
               list_data.append(j)
           b = 1
           for k,q in enumerate(list_data):
               for a in range(0,len(list_data[k])):
                   sheet1.write(b, a, list_data[k][a])
               b = b + 1
           datestring = datetime.strftime(datetime.now(), ' %Y-%m-%d %H-%M-%S')
           filename = tkinter.filedialog.asksaveasfilename(filetypes=[('xlsx', '*.xlsx')], initialfile="罪犯余额检测"+datestring, initialdir='C:\\')
           filename = filename + '.xls'
           # messagebox.showinfo("提示", "保存完毕~！！！")
           # file.save('E:\\lecent_all\\'+'罪犯余额检测'+datestring+'.xlsx')
           file.save(filename)

        Button(top, text='导出数据', command=report_data, fg='green').place(x=600, y=355)

    def zhzf_bank_check():
        capital_msg.place_forget()
        bank_msg.place_forget()
        balance_msg.pack_forget()
        zfzfh_bank_msg.pack()
        zfzfh_bank_msg.place(x=2, y=120)
        zfzfh_bank_msg.heading('0', text='姓名')
        zfzfh_bank_msg.heading('1', text='账号')
        zfzfh_bank_msg.heading('2', text='账户余额')
        zfzfh_bank_msg.heading('3', text='银行流水余额')
        zfzfh_bank_msg.heading('4', text='差额(账户表-流水表)')
        zfzfh_bank_msg.heading('5', text='账号状态')
        zfzfh_bank_msg.column('0', width=80, anchor='center')
        zfzfh_bank_msg.column('2', width=70, anchor='center')
        zfzfh_bank_msg.column('3', width=90, anchor='center')
        zfzfh_bank_msg.column('4', width=120, anchor='center')
        zfzfh_bank_msg.column('5', width=80, anchor='center')
        zfzfh_bank_msg_sql = "select * from (select (select xm from base_offender_info boi where boi.id=c.offender_id) as xm,c.bank_card_number,c.bank_balance,d.balance,c.bank_balance - d.balance as ce,\
                                (case (select status from base_offender_info boi where boi.id=c.offender_id) when 1 then '在狱'else '出狱' end )as status from base_offender_account c,\
                                (select * from bank_detailed_info a where exists (select * from (select t.cust_acc_num,max(t.transaction_date) as t_date,max(t.detailed_num) as mnum \
                                from bank_detailed_info t where length(t.cust_acc_num) = 19 and t.acc_detailed_type=1 group by cust_acc_num) b where a.cust_acc_num = b.cust_acc_num \
                                and a.transaction_date = b.t_date and b.mnum = a.detailed_num)) d where c.bank_card_number = d.cust_acc_num and c.bank_balance != d.balance) order by ce desc"
        zfzfh_bank_msg_sql_results = go(zfzfh_bank_msg_sql)
        x = zfzfh_bank_msg.get_children()
        for item in x:
            zfzfh_bank_msg.delete(item)
        if len(zfzfh_bank_msg_sql_results) > 0:
            capital_count.set(len(zfzfh_bank_msg_sql_results))
            j = 0
            for i in zfzfh_bank_msg_sql_results:
                zfzfh_bank_msg.insert('', 'end', values=zfzfh_bank_msg_sql_results[j])
                j = j + 1
        else:
            zfzfh_bank_msg.insert('', 'end', values='暂无结果')
            capital_count.set('0')

        # 导出btn
        def report_data():
            file = xlwt.Workbook()
            sheet1 = file.add_sheet('sheet1', cell_overwrite_ok=True)
            list_data = []
            for i in range(0, len(capital_title)):
                sheet1.write(0, i, zfzfh_bank_msg.heading(i)['text'])
            for j in zfzfh_bank_msg_sql_results:
                list_data.append(j)
            b = 1
            for k, q in enumerate(list_data):
                for a in range(0, len(list_data[k])):
                    sheet1.write(b, a, list_data[k][a])
                b = b + 1
            datestring = datetime.strftime(datetime.now(), ' %Y-%m-%d %H-%M-%S')
            filename = tkinter.filedialog.asksaveasfilename(filetypes=[('xlsx', '*.xlsx')],
                                                            initialfile="智慧政法银行余额检测" + datestring, initialdir='C:\\')
            filename = filename + '.xls'
            # messagebox.showinfo("提示", "保存完毕~！！！")
            # file.save('E:\\lecent_all\\'+'罪犯余额检测'+datestring+'.xlsx')
            file.save(filename)

        Button(top, text='导出数据', command=report_data, fg='green').place(x=600, y=355)

    Button(top, text='银行余额检测', command=capital_check,fg='green').place(x=2, y=80)
    Button(top, text='资金余额检测', command=bank_check, fg='green').place(x=100, y=80)
    Button(top, text='罪犯余额检测', command=balance_check, fg='green').place(x=200, y=80)
    Button(top, text='智慧政法1.0检测', command=zhzf_bank_check, fg='green').place(x=300, y=80)

#销售检测，资金明细表 减 实际销售金额
def create_newwin_sale():
    top = Toplevel()
    top.title('销售资金检测')
    top.geometry("850x400")
    top.resizable(0, 0)  # 锁定窗口
    sale_title = ("0", "1", "2", "3", "4", "5","6","7")
    sale_msg = ttk.Treeview(top, height=10, show="headings", columns=sale_title)  # capital_check表格
    # sale_msg.place(x=5,y=10)
    sale_msg.pack()

    sale_msg.heading('0',text="罪犯ID")
    sale_msg.heading('1', text="罪犯姓名")
    sale_msg.heading('2', text="银行账号")
    sale_msg.heading('3', text="实际销售金额")
    sale_msg.heading('4', text="消费结算金额")
    sale_msg.heading('5', text="差额(消费结算-实际销售)")
    sale_msg.heading('6', text="消费未结算")
    sale_msg.heading('7', text="账号状态")

    sale_msg.column('0',width=50,anchor='center')
    sale_msg.column('1', width=80, anchor='center')
    sale_msg.column('3', width=90, anchor='center')
    sale_msg.column('4', width=90, anchor='center')
    sale_msg.column('5', width=150, anchor='center')
    sale_msg.column('6', width=80, anchor='center')
    sale_msg.column('7', width=80, anchor='center')

    sale_total = tkinter.Label(top, text="数量：", font=("隶书", 10), fg="green")
    sale_total.place(x=150, y=240, anchor='nw')
    sale_count = StringVar()
    sale_counts = tkinter.Label(top, textvariable=sale_count, font=("隶书", 10), fg="green")
    sale_counts.place(x=200, y=240, anchor='nw')


    sale_data_search = "select  c.offender_id,c.offender_name,c.bank_card_number,c.实际销售金额,c.transaction_money,boa.consumption_money,c.chae,c.状态 from (select offender_id,offender_name,bank_card_number,实际销售金额,transaction_money,chae,状态 from (select a.xm,b.bank_card_number,\
                        a.transaction_money,b.实际销售金额,(a.transaction_money - b.实际销售金额) as chae,\
                       (case (select status from base_offender_info boi where boi.id = a.offender_id) when 0 then '出狱' else '在狱' end) as 状态,b.offender_name,b.offender_id \
                       from (select * from (select ci.offender_id,(select boa.bank_card_number from base_offender_account boa \
                        where boa.offender_id = ci.offender_id) as bank_card_number,(select xm from base_offender_info boi where boi.id = ci.offender_id) as offender_name,\
                        sum(case invoicing_type when 3 then (ci.total_selling_amount) when 4 then -(ci.total_selling_amount) else 0 end) 实际销售金额 from commodity_invoicing ci group by ci.offender_id) \
                        where offender_id is not null) b left join (select cdi.offender_id, \
                        (select boi.xm from base_offender_info boi where boi.id = cdi.offender_id) as xm, \
                        abs(sum(case cdi.transaction_type when 0 then cdi.transaction_money else -cdi.transaction_money end)) as transaction_money  \
                        from capital_detailed_info cdi where cdi.bill_number like 'XFJS%' \
                        group by cdi.offender_id) a on b.offender_id = a.offender_id order by (a.transaction_money - b.实际销售金额) asc) where chae != 0 or chae is null ) c  left join base_offender_account boa on c.offender_id = boa.offender_id and c.bank_card_number=boa.bank_card_number \
                        where chae + boa.consumption_money <> 0 or chae is null order by chae asc"
    sale_data_rul = go(sale_data_search)
    x = sale_msg.get_children()
    for item in x:
        sale_msg.delete(item)
    if len(sale_data_rul) > 0:
        sale_count.set(len(sale_data_rul))
        j = 0
        for i in sale_data_rul:
            sale_msg.insert('', 'end', values=sale_data_rul[j])
            j = j + 1
    else:
        sale_msg.insert('', 'end', values='暂无结果')
        sale_count.set('0')

    # 导出btn
    def report_data():
        file = xlwt.Workbook()
        sheet1 = file.add_sheet('sheet1', cell_overwrite_ok=True)
        list_data = []
        for i in range(0, len(sale_title)):
            sheet1.write(0, i, sale_msg.heading(i)['text'])
        for j in sale_data_rul:
            list_data.append(j)
        b = 1
        for k, q in enumerate(list_data):
            for a in range(0, len(list_data[k])):
                sheet1.write(b, a, list_data[k][a])
            b = b + 1
        datestring = datetime.strftime(datetime.now(), ' %Y-%m-%d %H-%M-%S')
        filename = tkinter.filedialog.asksaveasfilename(filetypes=[('xlsx', '*.xlsx')],initialfile="销售资金检测" + datestring, initialdir='C:\\')
        filename = filename + '.xls'
        file.save(filename)

    Button(top, text='导出数据', command=report_data, fg='green').place(x=10, y=240)
    sale_info = tkinter.Label(top, text="有差额的数据请核实该罪犯狱内消费数据！！1、请查看罪犯消费未结算和当月消费，以及是否调过账！！", font=("隶书", 10), fg="red")
    sale_info.place(x=10,y=280)


#一下一级窗口
comvalue=tkinter.StringVar()#窗体自带的文本，新建一个值
comboxlist=ttk.Combobox(win,width=65, height=20,font=1,textvariable=comvalue) #初始化
comboxlist["values"]=msg
comboxlist.current() #默认选择第一个
# comboxlist.bind("<<ComboboxSelected>>",go) #绑定事件,(下拉列表框被选中时，绑定go()函数)
comboxlist.place(x=10,y=40,anchor='nw')

#功能----上下账模块，接见款
w2 = tkinter.Label(win, text="1、检测：明细状态为0,2,9的；2、上下账处理：处理银行交易成功且未同步的明细的；3、状态修改：处理银行未交易且没有中间表的单子为待提交（含接见款）",font=("隶书",9), fg="red")
w2.place(x=10, y=130, anchor='nw')

button1 = tkinter.Button(win, text="上下账单检测", font=("隶书",10), command=prisoner_money_check,width=12, height=2, fg="green")
button1.place(x=10, y=90, anchor='nw')

button2 = tkinter.Button(win, text="上下账单处理", font=("隶书",10), command=create_newwin,width=12, height=2, fg="green")
button2.place(x=115, y=90, anchor='nw')

button3 = tkinter.Button(win, text="上下账处理中状态修改", font=("隶书",10), command=change_status_wid,width=19, height=2, fg="green")
button3.place(x=220, y=90, anchor='nw')

button4 = tkinter.Button(win, text="接见款处理中", font=("隶书",10), command=create_newwin_jjk, width=12, height=2, fg="green")
button4.place(x=370, y=90, anchor='nw')

button5 = tkinter.Button(win, text="罪犯编号修改", font=("隶书",10), command=change_prisoner_bh,width=12,height=2, fg="green")
button5.place(x=480, y=90, anchor='nw')

button9 = tkinter.Button(win, text="检测账户资金", font=("隶书",10), command=capital_test,width=12,height=2, fg="green")
button9.place(x=580, y=90, anchor='nw')

b1 = tkinter.Label(win, text="返回信息：",font=("隶书",10), fg="green")
b1.place(x=5, y=200, anchor='nw')

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

button7 = tkinter.Button(win, text="日志清除&空间查询", command=clear_catalina,font=("隶书",10), width=18, height=2, fg="green")
button7.place(x=10, y=300, anchor='nw')

#数据重复
button8 = tkinter.Button(win, text="银行流水重复删除", command=create_newwin_bankdet,font=("隶书",10), width=15, height=2, fg="green")
button8.place(x=170, y=300, anchor='nw')

button6 = tkinter.Button(win, text="销售数据检测", command=create_newwin_sale,font=("隶书",10), width=14, height=2, fg="green")
button6.place(x=300, y=300, anchor='nw')

#功能说明
tkinter.Label(win, text="以上功能已全部开放(内网直连，穿透不能使用)",font=("隶书",10), fg="green").place(x=10,y=360)

win.mainloop()