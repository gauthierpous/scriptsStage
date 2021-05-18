# -*- coding: utf-8 -*-
"""
Created on Tue May  4 12:17:29 2021

@author: 4165306
"""



import tkinter as tk

def quit(root):
    root.destroy()

root = tk.Tk()
tk.Button(root, text="Quit", command=lambda root=root:quit(root)).pack()
root.mainloop()

