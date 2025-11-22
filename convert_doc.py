import os
from win32com import client as wc


#
def convert_doc_to_docx(input_path, file_name, output_path=None):
    """
    单个doc文件转docx
    :param input_path: doc文件路径（自动识别该路径下的doc文件），路径参数须传相对路径
    :param file_name: 文件名称
    :param output_path: docx文件的输出路径（可选，如不填则和input路径相同）
    :return:
    """
    absolute_path = os.getcwd()
    doc_file_path = os.path.join(absolute_path, input_path, file_name)
    print("\n文件转换格式中 - {}".format(doc_file_path))
    word = wc.Dispatch("Word.Application")
    # raise ValueError('test error')
    doc = word.Documents.Open(doc_file_path)  # 打开文件
    if output_path is None:
        docx_file_path = os.path.join(absolute_path, input_path, '{}x'.format(file_name))
    else:
        docx_file_path = os.path.join(absolute_path, output_path, '{}x'.format(file_name))
    doc.SaveAs(docx_file_path, 12)  # 将文件另存为.docx。12表示docx格式
    doc.Close()
    word.Quit()
    
    print("格式转换完成，输出到 {}".format(docx_file_path))
    return docx_file_path

def batch_convert(input_path, output_path=None):
    """
    多个doc文件转docx
    :param input_path: doc文件路径（自动识别该路径下的doc文件），路径参数须传相对路径
    :param output_path: docx文件的输出路径（可选，如不填则和input路径相同）
    :return:
    """
    fail_file_list = []
    for file in os.listdir(input_path):
        # 找出文件中以.doc结尾并且不以~$开头的文件（~$是为了排除临时文件）
        if file.endswith('.doc') and not file.startswith('~$'):
            try:
                convert_doc_to_docx(input_path, file, output_path)
            except:
                fail_file_list.append(file)
                print('文件{}转换失败，跳过该文件'.format(file))
    print('\n转换失败文件列表：')
    for f in fail_file_list:
        print(f)


if __name__ == '__main__':
    input_path = "input"  # 输入相对路径（内为doc文件）
    output_path = "output"  # 输入相对路径
    batch_convert(input_path, output_path)
