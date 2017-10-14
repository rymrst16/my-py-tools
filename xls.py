#!/usr/bin/env python3
# -*- coding: utf-8 -*-
'''
本模块用来对excel表格进行操作,可以打开xls,xlsx,csv文件，也可以打开纯文本文件，纯文本文件每一行占一个数组中一列，形成一个n*1数组，
'''
import numpy as np
import pandas as pd
import math,xlrd,xlwt,openpyxl,csv,re
from xlutils.copy import copy as xlcopy
from pandas import Series,DataFrame

'''
def xlsread(file=None,worksheet=0,filter=True,number=True):
  
    #注意:本程序对于excel表格中为错误值的单元格读入时还是会显示数值，原因不明，暂时弃用
    本函数代码用来实现对excel表格的读取  
    参数：
    file参数：选择要打开的工作簿的路径
    worksheet参数：用来要读取的工作簿中的工作表，可以使用工作表名或者工作表的位置进行选择
    filter参数：用来设置是否需要过滤掉读取数据中的所有值都为空的行和列，csv文件只能得到过滤后的数据，该选项无效
    num参数：用来设置返回的数组是一个数值型的还是文本型的，默认为数值型的，数值型的数组会将表中的文本位置设为空值。
    返回值：
    返回一个np.array对象
    suffix = file.split('.')[-1]
    if suffix == 'xls' or suffix == 'xlsx': #处理xls(x)文件
         wb = xlrd.open_workbook(file)
         if isinstance(worksheet,int):
             ws = wb.sheets()[worksheet]
         elif isinstance(worksheet,str):
             ws = wb.sheet_by_name(worksheet)
         wbdata = []
         for i in range(ws.nrows):
             rowdata = ws.row_values(i)
             wbdata.append(rowdata)
         if filter == False:     #不需要过滤
             if number:
                return np.array(str2num(wbdata))
             else:
                return np.array(wbdata)
         elif filter == True:    #需要过滤
             #先转置，处理列为空的前几列，再转置回来，对为空的前几行进行处理，再将结果以np.array形式返回
             count = 0
             wbdata = np.array(wbdata).T.tolist() #获取转置后的列表表达
             for rowdata in wbdata:
                 if rowdata.count('') == len(rowdata):
                     count += 1
                 else:
                     break
             for num in range(count):
                 del wbdata[0]
             count = 0
             wbdata = np.array(wbdata).T.tolist()
             for rowdata in wbdata:
                 if rowdata.count('') == len(rowdata):
                     count += 1
                 else:
                     break
             for num in range(count):
                 del wbdata[0]
             if number:
                return np.array(str2num(wbdata))
             else:
                return np.array(wbdata)

    else:   #处理csv文件和纯文本文件
        with open(file) as csvfile:
            read = csv.reader(csvfile)
            wbdata = []
            for rowdata in read:
                wbdata.append(rowdata)
            if number:
                return np.array(str2num(wbdata))
            else:
                return np.array(wbdata)
'''    
def xlsread(file=None,worksheet=0,filter=True,number=True,skiprows=0,skipcols=0,skip_footer=0, header=None):
    '''
    本函数代码包装了pandas的pd.read_excel方法，使得其功能类似于xlsread
    参数：
    file参数：选择要打开的工作簿的路径
    worksheet参数：用来要读取的工作簿中的工作表，可以使用工作表名或者工作表的位置进行选择
    filter参数：用来设置是否需要过滤掉读取数据中的所有值都为空的行和列，csv文件只能得到过滤后的数据，该选项无效
    num参数：用来设置返回的数组是一个数值型的还是文本型的，默认为数值型的，数值型的数组会将表中的文本位置设为空值。
    skiprows参数：用来选择过滤掉所读取文本的开头多少行
    skip_footer参数：用来选择过滤掉所读取文本的后尾的多少行
    header参数：用来设置选择第几行作为pandas的标签，一般不用修改
    返回值：
    返回一个np.array对象
    '''   
    pdxlsdata = pd.read_excel(file,worksheet,None,skiprows,skip_footer)
    pdxlsdata = pdxlsdata.ix[:,skipcols:]
    wbdata = np.array(pdxlsdata).tolist()
    if filter == False:     #不需要过滤
        if number:
            return np.array(str2num(wbdata))
        else:
            return np.array(num2str(wbdata))
    elif filter == True:    #需要过滤
        #先转置，处理列为空的前几列，再转置回来，对为空的前几行进行处理，再将结果以np.array形式返回
        count = 0
        wbdata = np.array(wbdata).T
        wbrow = wbdata.shape[0]
        for rowdata in range(wbrow):
            if np.count_nonzero(wbdata[rowdata] != wbdata[rowdata]) == len(wbdata[rowdata]):
                count += 1
            else:
                break
        wbdata = wbdata.tolist() 
        for num in range(count):
            del wbdata[0]
        count = 0
        wbdata = np.array(wbdata).T
        wbrow = wbdata.shape[0]
        for rowdata in range(wbrow):
            if np.count_nonzero(wbdata[rowdata] != wbdata[rowdata]) == len(wbdata[rowdata]):
                count += 1
            else:
                break
        wbdata = wbdata.tolist()         
        for num in range(count):
            del wbdata[0]
        if number:
             return np.array(str2num(wbdata))
        else:
             return np.array(wbdata)

def xlswrite(srcdata = None,file=None,worksheet=0,row=1,col=1,cover=False,formatting=False):
    '''
    本函数代码用来实现将数据写入excel表格，其功能类似于xlswrite，如果file路径的文件不存在，则会创建一个新的工作簿，如果存在，则通过cover参数来选择是否覆盖已经存在的工作簿，默认情况下是不覆盖的（False），row和col参数用于选择需要填写的数据的第一个元素的位置，即左上角的位置写入哪里，空缺值在表格中表现为空白。formatting参数为不覆盖时，即续写时，保留原来的格式，不过只对.xls文件有效，.xlsx文件暂时不支持。
    参数：
    srcdata参数：传入的是一个np.array
    file参数：选择目标工作簿的路径
    worksheet参数：用来选择将数据写入工作簿的第几张工作表，默认为第一张（下标为0的工作表）
    row，col参数：分别表示将数据写入的第一个单元格（即左上角的单元格的位置），从1开始计数。
    cover参数：选择是否覆盖已经存在的工作簿，默认情况下是不覆盖的（False）
    formatting参数：续写时，保留原来的格式，不过只对.xls文件有效，.xlsx文件暂时不支持,默认不保留原来的格式。
    返回值：
    返回一个np.array对象
    '''
    if len(srcdata.shape) == 2:
        srcdata = srcdata.tolist()
    elif len(srcdata.shape) == 1:
        srcdata = np.array(np.matrix(srcdata)).tolist()
    suffix = file.split('.')[-1] #获取后缀，分别对.xls和.xlsx文件进行处理
    if cover == True:   #覆盖原文件
        if suffix == 'xls':
            wb = xlwt.Workbook()
            if isinstance(worksheet,str):
                ws = wb.add_sheet(worksheet)
            else:
                for i in range(0,worksheet+1):
                    ws = wb.add_sheet('Sheet'+str(i+1))      
            (m,n) = (len(srcdata),len(srcdata[0]))
            for rowposition in range(m):
                for colposition in range(n):
                    ws.write(row-1+rowposition,col-1+colposition,srcdata[rowposition][colposition])
            wb.save(file)
        elif suffix == 'xlsx':
            wb = openpyxl.Workbook()
            if isinstance(worksheet,str):
                ws = wb.create_sheet(title=worksheet)
            else:
                for i in range(0,worksheet+1):
                    ws = wb.create_sheet(title=('Sheet'+str(i+1)))
            wb.remove_sheet(wb.get_sheet_by_name("Sheet"))
            (m,n) = (len(srcdata),len(srcdata[0]))
            for rowposition in range(m):
                for colposition in range(n):
                    ws.cell(row=(row+rowposition),column=col+colposition).value=srcdata[rowposition][colposition]
            wb.save(filename=file)
    elif cover == False:    #不覆盖原文件
        if suffix == 'xls':    #处理后缀为.xls文件，主要是格式方面是否保留
            try:    #是否存在该工作簿，如果存在则在该工作簿上面续写，否则创建一个同名工作簿写入
                wbr= xlrd.open_workbook(file,formatting_info=formatting)
                wbw = xlcopy(wbr)
                try:
                    ws = wbw.get_sheet(worksheet) 
                except IndexError:  #若用角标取工作表，如果该下标的工作表不存在，则做如下处理，即创建一个默认名称工作表，写入数据
                    wbnamelist = wbr._sheet_names
                    candlist = [sinname for sinname in wbnamelist if re.match(r'^[Ss]heet[0-9]+$', sinname)]
                    if candlist:
                        candsheetnum = [int(getsheetnum.split('t')[-1]) for getsheetnum in candlist]
                        for i in range(0,worksheet-len(wbnamelist)+1):
                            ws = wbw.add_sheet(('Sheet'+str(max(candsheetnum)+1+i)))
                    else:
                        ws = wbw.add_sheet('Sheet1') 
                except Exception:   #若用工作表名称取工作表，如果该名称的工作表不存在，则做如下处理，即创建一张以该名称命名的工作表，并将数据写入
                    ws = wbw.add_sheet(worksheet) 
                (m,n) = (len(srcdata),len(srcdata[0]))
                for rowposition in range(m):
                    for colposition in range(n):
                        ws.write(row-1+rowposition,col-1+colposition,srcdata[rowposition][colposition])
                wbw.save(file)
            except FileNotFoundError:   #不存在该工作簿时，创建同名工作簿写入
                wb = xlwt.Workbook()
                if isinstance(worksheet,str):
                    ws = wb.add_sheet(worksheet)
                else:
                    for i in range(0,worksheet+1):
                        ws = wb.add_sheet('Sheet'+str(i+1))      
                (m,n) = (len(srcdata),len(srcdata[0]))
                for rowposition in range(m):
                    for colposition in range(n):
                        ws.write(row-1+rowposition,col-1+colposition,srcdata[rowposition][colposition])
                wb.save(file)
        elif suffix == 'xlsx':
            try:    #是否存在该工作簿，如果存在则在该工作簿上面续写，否则创建一个同名工作簿写入
                wb = openpyxl.load_workbook(file)
                try:
                    if isinstance(worksheet,str):
                        ws = wb.get_sheet_by_name(worksheet)
                    else:
                        ws = wb.get_sheet_by_name(wb.sheetnames[worksheet])
                except IndexError:  #若用角标取工作表，如果该下标的工作表不存在，则做如下处理，即创建一个默认名称工作表，写入数据
                    wbnamelist = wb.sheetnames
                    candlist = [sinname for sinname in wbnamelist if re.match(r'^[Ss]heet[0-9]+$', sinname)]
                    if candlist:
                        candsheetnum = [int(getsheetnum.split('t')[-1]) for getsheetnum in candlist]
                        for i in range(0,worksheet-len(wbnamelist)+1):
                            ws = wb.create_sheet(title=('Sheet'+str(max(candsheetnum)+1+i)))
                    else:
                        ws = wbw.add_sheet('Sheet1') 
                except Exception:   #若用工作表名称取工作表，如果该名称的工作表不存在，则做如下处理，即创建一张以该名称命名的工作表，并将数据写入
                    ws = wb.create_sheet(title=worksheet) 
                (m,n) = (len(srcdata),len(srcdata[0]))
                for rowposition in range(m):
                    for colposition in range(n):
                        ws.cell(row=(row+rowposition),column=col+colposition).value = srcdata[rowposition][colposition]
                wb.save(file)
            except FileNotFoundError:   #不存在该工作簿时，创建同名工作簿写入
                wb = openpyxl.Workbook()
                if isinstance(worksheet,str):
                    ws = wb.create_sheet(title=worksheet)
                else:
                    for i in range(0,worksheet+1):
                        ws = wb.create_sheet(title=('Sheet'+str(i+1)))
                wb.remove_sheet(wb.get_sheet_by_name("Sheet"))
                (m,n) = (len(srcdata),len(srcdata[0]))
                for rowposition in range(m):
                    for colposition in range(n):
                        ws.cell(row=(row+rowposition),column=col+colposition).value=srcdata[rowposition][colposition]
                wb.save(filename=file)

            
def isnum(num):
    '''
    用来判断num的值是否为数值型，是的话返回True
    '''
    try:
        float(num)
        return True
    except ValueError:
        return False

def str2num(strlist):
    '''
    用来将数组中的字符串元素转为数组元素，非数值的数据填充为nan。
    '''
    m = len(strlist) ; n = len(strlist[0])
    for row in range(m):
        for col in range(n):
            if isnum(strlist[row][col]):
                strlist[row][col] = float(strlist[row][col])
            else:
                strlist[row][col] = np.nan
    return strlist

def num2str(strlist):
    '''
    用来将数组中的字符串元素转为数组元素，非数值的数据填充为nan。
    '''
    m = len(strlist) ; n = len(strlist[0])
    for row in range(m):
        for col in range(n):
            if math.isnan(strlist[row][col]):
                strlist[row][col] = np.nan
            elif isnum(strlist[row][col]):
                strlist[row][col] = str(strlist[row][col])
            else:
                strlist[row][col] = strlist[row][col]
    return strlist

    
    
        
        
    
    
