#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
本程序用来对数据进行存储和读取
"""
import pickle

def save(filename,**arr):
    '''
    本函数用来存储内存中的变量，其中第一个参数是文件名，第二个参数是关键字参数，用来表示要存储的变量，如要存储s和a则填s=s，a=a。
    这里其实是将变量和变量值的字典序列化存储了。
    '''
    pickle.dump(arr,open(filename,'wb'))

def load(filename):
    '''
    本函数用来读取存储了变量的文件，将变量以字典的形式返回，在交互式环境下可以使用locals().update(data)，将字典转为当前工作空间的变量
    '''
    data = pickle.load(open(filename,'rb'))
    return data
    
    
