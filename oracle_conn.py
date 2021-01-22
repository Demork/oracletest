#!/usr/bin/env python
# -*- coding: utf-8 -*-
import cx_Oracle as oea
import sys
import os

# 字符转中文防止中文乱码
os.environ['NLS_LANG'] = 'SIMPLIFIED CHINESE_CHINA.UTF8'


# 定义变量值
# IP地址，用户名，密码，数据库名，字符格式
def admin(self):
    return cx_Oracle.connect(self.user, self.pwd, '{0}:{1}/{2}'.format(self.host, self.port, self.sid))


o_ip_addr = '192.168.1.25'
o_user_name = "lecent_sq"
o_user_psd = 'lecent_sq'
o_port = '1521'
o_database = 'pora12c1.lecent.domain'

try:  # 尝试连接数据库
    # 这里的顺序是用户名/密码@oracleserver的ip地址/数据库名字,多种连接方式
    # conn = cx_Oracle.connect('lecent_dx', 'lecent_dx', '192.168.1.25:1521/pora12c1.lecent.domain')  # 定义连接数据库的信息
    conn = oea.connect(o_user_name, o_user_psd, o_ip_addr + ':' + o_port + '/' + o_database)
# conn = cx_Oracle.connect("lecent_dx/lecent_dx@192.168.1.25:1521/pora12c1.lecent.domain")
    print(o_user_name, o_user_psd, o_ip_addr + ':' + o_port + '/' + o_database)
except cx_Oracle.OperationalError as message:  # 连接失败提示
    print('数据库连接失败！请检查数据库连接的IP、用户名、密码、数据库名和端口是否正确。')

cur = conn.cursor()  # 定义连接对象
'''
sql = "select ci.offender_id 罪犯ID,ci.offender_name 罪犯姓名," \
      "sum(case invoicing_type when 3 then (ci.total_selling_amount) when 4 then - " \
      "(ci.total_selling_amount) else 0 end) 实际销售金额 from commodity_invoicing ci " \
      "where ci.create_date >= to_date('2019-11-01', 'yyyy-mm-dd') and ci.offender_id " \
      "is not null group by ci.offender_id, offender_name"
      '''
sql = "select * from base_offender_info where 1=1"
result = cur.execute(sql).fetchall()
# data = result.fetchall()
# cursor.execute() 使用cursor提供的方法来执行查询语句
# cursor.fetchall()  # 使用fetchall方法返回所有查询结果

# for Result in data:  # 将结果集遍历打印

title = [i[0] for i in cur.description]
print(title)

for Result in result:
    print(Result)  # 打印查询结果
cur.close()  # 关闭cursor对象
# conn.commit()
conn.close()  # 关闭数据库链接
