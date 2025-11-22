# -*- coding: utf-8 -*-
"""
Created on Sun Apr 20 10:54:34 2025

@author: jinn
"""
import DocRead as dr
import numpy as np
import pandas as pd
from docx import Document
import ArrayCmp as ac
#import TransferDataFile as tdf
import filemanage as fm
import os

import os

# 获取当前目录下的word文件夹路径    
path = os.path.join(os.getcwd(), 'word') + os.sep
filename = '协议模板（公开）.docx'
fileAbsPath = os.path.join(path, filename)

# 把输出文件放在同一文件夹中
docxname = fm.transfer_docx_and_doc(filename, path, path)
if docxname is None:
    print(docxname)
else:    
    outputpath = os.path.join(path, 'csvfile') + os.sep
    #filenamedocx = filename.replace('doc', 'docx')
    #docxname = path + filenamedocx
    #splitname = filename.split(".")
    
    print("filename is ",docxname)
    doc = Document(docxname)
    fm.remove_directory(outputpath)
    # #消息名称头部关键字段
    # msgHeadNameArray = ["信息名称","信息标识"]
    # baseline = ["代号","内容","类型","字节","值域","数据处理","单位","备注"]
    # appointIndexArray = [1,2,6,7]
    # indexCtrlArray = [2,0,4,5]
    
    # msgHeadNameArray = ["信息名称","信息标识"]
    # baseline = ["代号","内容","类型","字节","值域","数据处理","单位","区间","备注"]
    # appointIndexArray = [1,2,6,8]
    # indexCtrlArray = [2,0,4,5]
    msgHeadNameArray = ["信息名称","上级信息名称"]
    baseline = ["序号","参数","数据类型","数据长度（字节）","值域","单位","备注"]
    appointIndexArray = [1,2,4,6]
    indexCtrlArray = [1,0,1,2]
    # msgHeadNameArray = ["信息名称","上级信息名称"]
    # baseline = ["序号","代号","数据含义","数据类型","字节数","取值范围","备注"]
    # appointIndexArray = [2,3,5,6]
    # indexCtrlArray = [7,0,1,2]
    
    #msgHeadNameArray = ["信息名称","上级信息名称"]
    # baseline = ["序号","数据含义","数据类型","字节数","取值范围","备注"]
    # appointIndexArray = [1,2,4,5]
    # indexCtrlArray = [7,0,1,2]
    # msgHeadNameArray = ["信息名称","上级信息名称"]
    # baseline = ["序号","名称","数据类型","字节数","备注"]
    # appointIndexArray = [1,2,3,4]
    # indexCtrlArray = [2,0,1,2]

    # msgHeadNameArray = ["信息名称","上级信息名称"]
    # baseline = ["序号","名称","数据类型","单位","备注"]
    # appointIndexArray = [1,2,3,4]
    # indexCtrlArray = [2,0,1,2]
# =============================================================================
#     msgHeadNameArray = ["信息名称","上级信息名称"]
#     baseline = ["序号","信号名称","数据类型","字节数","说明"]
#     appointIndexArray = [1,2,4,4]
#     indexCtrlArray = [2,0,1,2]
# =============================================================================

# =============================================================================
#     msgHeadNameArray = ["信息名称","上级信息名称"]
#     baseline = ["序号","字段","类型","单位","值域","备注"]
#     appointIndexArray = [1,2,3,5]
#     indexCtrlArray = [2,0,1,2]
# 
# =============================================================================

# =============================================================================
#     msgHeadNameArray = ["信息名称","上级信息名称"]
#     baseline = ["序号","信号名称","数据类型","字节数","说明"]
#     appointIndexArray = [1,2,4,4]
#     indexCtrlArray = [2,0,1,2]
# =============================================================================
# =============================================================================
#     msgHeadNameArray = ["通信帧名称","信息流向"]
#     baseline = ["序号","内容","长度","值","单位","值域","数据转换方法","说明"]
#     appointIndexArray = [1,2,5,7]
#     indexCtrlArray = [2,0,3,9]
# =============================================================================

    """
        indexCtrlArray:list
        内容分为：
        startTableIndex : TYPE
            控制从哪个table开始找下一个table
        msgNameIndex:
            描述消息名称的行从第几行开始
        headIndex:
        数据消息头从第几行开始
        dataStartIndex：
            数据从第几行开始
    """

    
    dr.get_msg_data(doc, filename, outputpath,msgHeadNameArray,baseline,indexCtrlArray,appointIndexArray)