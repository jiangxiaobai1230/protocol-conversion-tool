# -*- coding: utf-8 -*-
"""
Created on Mon Aug 19 20:26:38 2024

@author: Administrator
"""
import os
import shutil
import convert_doc as cvd

def remove_directory(dir_path):
    if os.path.isdir(dir_path):
        shutil.rmtree(dir_path)
    os.mkdir(dir_path)  # 如果需要，可以在删除后重新创建目录
def return_file_list(filepath):
    """
    剔除非文件临时文件，返回文件列表

    Parameters
    ----------
    filepath : TYPE
        DESCRIPTION.

    Returns
    -------
    文件名list

    """ 
    filelist = []
    for file in os.listdir(filepath):
      file_path = os.path.join(filepath, file)
      if os.path.isfile(file_path) and (file.startswith('~') == False):
          #检查文件名称后缀，以及名称是否包含协议
          filelist.append(file)
    if filelist == []:
        print("输入文件夹内文件为空")
    return filelist
          
def transfer_docx_and_doc(file,outputpath,inputpath):
    """
    对file，如果是docx则剪切至outputpath,否则转换

    Parameters
    ----------
    file : TYPE
        DESCRIPTION.
    outputpath : TYPE
        DESCRIPTION.
    inputpath : TYPE
        DESCRIPTION.

    Returns
    -------
    None.

    """
    outputpath = ""
        #检查文件名称后缀，以及名称是否包含协议
    if ("协议" in file ):
        #如果是docx则转换，否则移动到output文件夹
        if file.endswith("docx"):
            destination_path = outputpath + file
            source_path = inputpath +  file
            os.rename(source_path, destination_path)
            outputpath = destination_path
        elif file.endswith("doc"):
            #调用convert_docx
            outputpath = cvd.convert_doc_to_docx(inputpath,file,outputpath)
            
        else:
            print(file,"文件类型超出范围，不进行转换")
    else:
        print(file,"文件没有协议字段不进行转换")
    return outputpath       
def transfer_protocolfile_type(filepath,outputpath,filename = None):
    """
    对filepath中的文件进行筛选，选出doc或docx的文件
    然后对doc文件进行转换
    复制docx文件
    输出在outputpath文件夹中

    Parameters
    ----------
    filepath : TYPE
        DESCRIPTION.

    Returns
    -------
    None.

    """
     #创建输出目录
    remove_directory(outputpath)
    if None != filename:
        transfer_docx_and_doc(filename,outputpath,filepath)
    else:
        
        filelist = return_file_list(filepath)
        print("移动并转换协议文件：")
        print(filelist)
        for file in filelist:
           transfer_docx_and_doc(file,outputpath,filepath)
                        
                    
                    
                
                
            