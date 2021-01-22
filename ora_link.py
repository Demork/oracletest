import os
import sys
import cx_Oracle as cx
db_port = '1521'
ser_name = 'pora12c1.lecent.domain'
class link_oracle:
    f = open("./oracle_link.txt", "r")
    lines = f.readlines()
    msg=[]
    for line in lines:
        line = line.rstrip("\n")
        msg.append(line)
    print(msg)





