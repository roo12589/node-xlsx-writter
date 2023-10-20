# -*- coding: utf-8 -*-
'''
@Time    :2023/8/29 17:27
@作者    :一晌小贪欢
@联系    :278865463@qq.com
'''

import win32com.client as win32
import os


def transform_xls(_input_path, _output_path):
    # 需要转换的文件路径
    input_path = _input_path

    # 转换完后输出的路径
    output_path = _output_path

    # 遍历需要转换的文件夹下面所有的文件
    file_list = os.listdir(input_path)

    # 获取遍历完的文件数量
    num = len(file_list)

    # 打印文件数量
    print(num)

    # 遍历文件
    for i in range(num):

        # 将文件和格式分开
        file_name = os.path.splitext(file_list[i])

        # 打印分开后的列表
        print(file_name)

        # 当遍历到的文件格式为'.xlsx'时
        if file_name[1] == '.xlsx':
            # 得到要转换的文件
            transfile1 = input_path + file_list[i]

            # 转换完需要输出的文件
            transfile2 = output_path + file_name[0]

            # 打印要转换的文件
            print('transfile1:' + str(transfile1))

            # 使用win32操作excel
            xlApp = win32.gencache.EnsureDispatch('Excel.Application')

            # 后台运行, 不显示，不警告
            # 不写这个会卡死……注意Python、win32需要保持一致。比如我的都是64位的
            xlApp.Visible = False
            xlApp.DisplayAlerts = False

            # 打开要转换的excel
            xls = xlApp.Workbooks.Open(transfile1)

            # 将需要转换的excel另存为xls格式。 56为xls
            xls.SaveAs(transfile2 + '.xls', FileFormat=56)

            # 关闭excel文件
            xls.Close()

            # 退出进程
            xlApp.Application.Quit()


if __name__ == '__main__':
    current_path = os.getcwd()
    
    # 待转换文件所在目录
    input_path = os.path.join(current_path,"./out/xlsx/")

    # 转换文件存放目录
    output_path = os.path.join(current_path, './out/')

    transform_xls(input_path, output_path)
