# -*- encoding:utf-8 -*-
from tkinter import *

root = Tk()

def create():
    top = Toplevel()
    top.title('Python')

    v1 = StringVar()
    e1 = Entry(top, textvariable=v1, width=10)
    e1.grid(row=1, column=0, padx=1, pady=1)

    Button(top, text='出现2级').grid(row=1, column=1, padx=1, pady=1)


Button(root, text='点击1级', command=create).pack()

mainloop()