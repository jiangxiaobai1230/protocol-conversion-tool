# -*- coding: utf-8 -*-
"""
Created on Sat May 25 09:50:55 2024
比较两个列表内容是否相等
@author: Administrator
"""

from collections import Counter
import copy
import re
import pandas as pd
import numpy as np
def are_lists_equal(list1, list2):
    return Counter(list1) == Counter(list2)
def compare_list_return_index(list1,list2):
    """
    比较两个list，
    以list1为基准，如果list2[i]与list1[i]不同则返回i
    两个列表的顺序一致，只是list2中有缺项
    Parameters
    ----------
    list1 : TYPE
        基准列表.
    list2 : TYPE
        待比较列表.

    Returns
    -------
    list，不同值的索引.

    """
    list_tmp = []
    "list2可能比list1短，检查list1中哪些元素list2没有，如果没有则返回list1的index"
    if len(list1) < len(list2):
        print("compare_list_return_index:list2维数更大")
    else:
        for index,tmp in enumerate(list1):
            if tmp in list2:
                pass
            else:
                list_tmp.append(index)
    return list_tmp

def list_copy(list1,list2):
    """
    将list1的值赋值给list2

    Parameters
    ----------
    list1 : TYPE
        赋值方
    list2 : TYPE
        被赋值.

    Returns
    -------
    None.

    """

    for strtmp in list1:
        list2.append(strtmp)
            
def get_unique_array(tmpList,uniqueStr):
    "去掉字符串中重复的元素,返回去重后的索引"
    "如tmplist = [1,2,2,3,4],则索引为 [0,1,,3,4]"
    baseStr = ""
    tmpIndex = 0
    indexArray = []
    for tmpStr in tmpList:
        if tmpStr == baseStr:
            pass
        else:
            uniqueStr.append(tmpStr)
            baseStr = tmpStr
            indexArray.append(tmpIndex)
        tmpIndex += 1
    return copy.deepcopy(indexArray)



def get_clean_chsArray(tmpList, cleanArray):
    """
    将tmpList中的换行符和空格去掉，仅保留中文和括号，并将英文括号替换为中文括号
    """
    cleanArray.clear()
    for tmpStr in tmpList:
        # 使用正则表达式保留中文字符和括号
        cleaned_str = re.sub(r'[^\u4e00-\u9fff()（）]', '', tmpStr)
        # 将中文括号替换为英文括号
        cleaned_str = cleaned_str.replace('(', '（').replace(')', '）')
        cleanArray.append(cleaned_str)

def transfer_shortname(longName):
    """
    读取longName，删除尾部的通信协议，从后往前数小于20字符的字符
    """
    tmpName = longName.replace('通信协议','')
    nameLenth = len(tmpName)
    shortName = ''
    if nameLenth>21:
        #从后往前取20字符
        shortName = tmpName[:20]
    else:
        shortName = tmpName
    
    return shortName


def transfer_data_type(list,translist,transLengthList):
    "对数据类型进行合法性检查，然后根据类型转换数据长度"
    #标准数据类型
    constTransDataType = ["INT8","UINT8","INT16","UINT16","INT32",
                          "UINT32" ,"FLOAT","DOUBLE"]
    constDataType = ["CHAR","UCHAR","SHORT","USHORT","INTERGER-32",
                     "UINTEGER-32","FLOAT","DOUBLE"]
    constDataLength = [8,8,16,16,32,32,32,32]
    for tmpStr in list:
        if tmpStr in constDataType:
            fstIndex = constDataType.index(tmpStr)
            translist.append(constTransDataType[fstIndex])
            transLengthList.append(constDataLength[fstIndex])
        else:
            print("new data type",tmpStr)
            translist.append(tmpStr)
            transLengthList.append(0)
            
def get_appoint_array(tmpList, listIndex):
    """
    "根据listIndex的索引值从tmplist中提取数据放入返回值"
    "listIndex的维数不应该超过tmplist"

    Parameters
    ----------
    tmpList : TYPE
        DESCRIPTION.
    listIndex : TYPE
        DESCRIPTION.
    appointArray : TYPE, optional
        根据索引跳出来的数组. The default is [].

    Returns
    -------
    None.

    """
    appointArray = []
    indexNum = len(listIndex)
    if indexNum > len(tmpList):
        print(" get_unique_array:输入数组维数异常")
    else:       
        i = 0
        while i < indexNum:
            appointArray.append(tmpList[listIndex[i]])
            i += 1
    return copy.deepcopy(appointArray)

def standard_datatype(oldType):
    """
    对数据类型字符转进行转换，包括两个工作
    1、ushort等转换为UINT16
    2、赋值一个bit序列，比如UINT16，bit序列为8
    3、如果没有找到转换方法则数据类型返回原值，bit类型为0

    Parameters
    ----------
    oldType : TYPE
        原数据类型.

    Returns
    -------
    转换后的数据类型和bit类型

    """ 
    
    "有无符号"
    symblePattern = r'u\w+'
    symbleMatch = None
    try:
        symbleMatch = re.match(symblePattern,oldType,re.I)
    except TypeError:
        print("standard_datatype:no matching",oldType)
        
    newType = ""
    bitType = ""
    if None == symbleMatch:
        "有符号"
        pass
    else:
        "无符号"
        newType += 'U'
        
    "类型及长度"
    #lengthPattern = ['char','short','integer','float','double']
    lengthPattern = ['CHAR','SHORT','INTEGER-32','FLOAT','DOUBLE']
    lengthstr = ['INT8','INT16','INT32','FLOAT32','DOUBLE64']
    bitStr = ['8','16','32','32','64']
    for index,tmpPattern in enumerate(lengthPattern):
        #比较是否两个类型都不满足
        try:
            if None == (re.search(tmpPattern,oldType,re.I)) and \
                None == (re.search(lengthstr[index],oldType,re.I)) :
                pass
            else:
                newType += lengthstr[index]
                bitType = bitStr[index]
                break
        except TypeError:
             print("standard_datatype:no matching",oldType)
             
    if newType == "" or bitType == "":
        print("standard_datatype:no matching",oldType)
        bitType = "0"
        newType = "Nan"
    return bitType,newType        
        
def get_bus_msg_param(addrStr):
    """
    将总线地址字符串拆分为信源、信宿、子地址、长度
    如BC RT2-SA1-7
    分为
    BC RT2 1 7
    如果是模式码则长度为0

    Parameters
    ----------
    addrStr : TYPE
        DESCRIPTION.

    Returns
    -------
    信源、信宿、子地址、长度

    """
    "信源,有时BC与RT没有空格，采用字符串头部匹配"
    xydata = ""
    xsdata = ""
    matchBC = re.match(r"BC",addrStr)
    rtStr = re.split(r"BC", addrStr)
    if None == matchBC:
        "开头不是BC"
        xsdata = "BC"
        splitStr = rtStr[0]
    else:
        xydata = "BC"
        splitStr = rtStr[1]
        
                  
    xslist = re.split(r"-",splitStr)
    rtcode = xslist[0]
    rtsubaddr = xslist[1][2:]
    if "模式码" in xslist[2]:
        dataLength = ["0"]
    else:
        dataLength = re.findall(r'\d+',xslist[2])
    if xydata == "":
        xydata = rtcode
    elif xsdata == "":
        xsdata = rtcode
    else:
        print("get_bus_msg_param:data set error")
    return xydata,xsdata,rtsubaddr,dataLength[0]
            
def extract_appoint_list(oldlist,indexarray,newlist):
    """
    根据indexarray提取oldlist中元素，组成一个新的list

    Parameters
    ----------
    oldlist : TYPE
        DESCRIPTION.
    indexarray : TYPE
        DESCRIPTION.

    Returns
    -------
    None.

    """
    for index in indexarray:
        newlist.append(oldlist[index])
def sort_dataframe_with_content(df,by,matelist):
    """
    在by索引上，依据matelist对df整行进行排序
    如果有匹配不上的字段则提示

    Parameters
    ----------
    df : dataframe
        被排序的数组.
    by : str
        列索引.
    matelist : list
        匹配字符.

    Returns
    -------
    排序后的dataframe.

    """
    "把by设置为行index"
    "先根据matelist对信息内容进行匹配赋值序号"
    "以matelist序号为准"
        
    emptyArray = np.zeros(len(df.index)) 
    dfindex =  len(df.index)

    df.loc[:,'sortnum'] = emptyArray
    emptyRow = np.zeros(len(df.columns))
    sortDF = pd.DataFrame(columns=df.columns.values.tolist())
    mateNum = 0

    i = 0
    while i <  dfindex:
        if matelist.count(df.at[i,by]) > 0:
            matchIndex = matelist.index(df.at[i,by])
            df.at[i,'sortnum'] = matchIndex
            mateNum += 1
        else:
            print("sort_dataframe_with_content:名称不匹配")
            df.at[i,'sortnum'] = 0xFF
        i += 1
    df = df.sort_values(by = "sortnum",ignore_index = True)  
    "删除匹配不上的行"
    df = df.drop(np.arange(mateNum,dfindex))     
    "对序号不连续的填充空ID"
    if mateNum <=  dfindex:
        "如果有匹配不上的，需要在df[by]中不匹配的地方加空格"
        "遍历序号的连续性，在不连续的地方填充"      
        i = 0
        while i < (len(df.index)):
            tmpdf = df.loc[i].to_frame().transpose()
            sortDF = pd.concat([sortDF,tmpdf],ignore_index=True)
            if (i+1) >= len(df.index):
                break
            else:
                setnum = df.at[i+1,"sortnum"]-df.at[i,"sortnum"]-1
                j = 0
                while j < setnum:
                    sortDF.loc[len(sortDF.index)] = emptyRow
                    j += 1 
 
            i += 1
    else:
        pass
     
    sortDF = sortDF.drop('sortnum',axis = 1)
    return sortDF

    
            
def extract_proto_type(str):
    "根据str中的关键字区分协议类型"
    if "总线通信协议" in str:
        print("该协议是总线通信协议")
        return 0
    elif "串行通信协议" in str:
        print("该协议是串行通信协议")
        return 1
    elif "网络通信协议" in str:
        print("该协议是网络通信协议")
        return 2
    else:
        print("无法定义该协议类型")
        return -1