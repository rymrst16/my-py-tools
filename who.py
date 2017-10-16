#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
用来显示当前命名空间中的所有变量（可以在调试中使用），使用的方法和matlab的who和whos类似
使用方法：
from who import *
whos():返回变量和值得字典
who(whos()):返回当前作用域中所有变量名（局部变量）
"""
whos = locals
def who(whos):
    print(list(whos.keys()))