#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
本程序用来对数据库进行操作，包括连接数据库，以及关闭连接
"""
#host='localhost';user='root';port=3306;passwd='123456';db='test';charset='utf8'
#(connect,cursor) = mydb.condb('localhost',3306,'root','123456','test')
import pymysql
def condb(host,port,user,passwd,db,charset='utf8'):
    '''
    本函数代码用来实现对指定数据库的连接，并返回游标。
    参数：
    host参数：选择主机名
    post参数：选择端口号
    user参数：填写用户名
    ppasswd参数：填写用户名的密码
    db参数：填写要连接的数据库名字
    charset参数：选择字符集
    返回值：
    返回一个游标，用来操作该数据库
    '''   
    connect = pymysql.Connect(host=host,port=port,user=user,passwd=passwd,db=db,charset=charset)    #创建数据库连接对象，用来连接数据库。这里注意：默认情况下pymysql.Connect方法中的charset值为latin-1，此时会对中文 不支持。
    #获取游标
    try:
        cursor = connect.cursor()
    except:
        connect.close()
    return (connect,cursor)

def closedb(connect,cursor):
    '''
    用于断开和数据库的连接，释放资源
    '''
    cursor.close()
    connect.close()
    

    

