'''
最基本的UI自动化测试脚本
此脚本step,method,selector为必填字段，若method为fill(输入),则value为必填
只需在testcase.xls文件中填入用例标题，操作步骤，操作方式，元素定位，输入值，断言元素定位，断言文本，运行run.py文件即可自动执行自动化测试
目前只支持文本断言，需填入断言元素定位和断言文本
执行结果会自动在result列填入pass/false
失败操作会在同目录自动创建screenshot文件夹放截图

关联三方包
pip install xlutils
pip install xlrd
pip install xlwt
pip install playwright
playwright install
'''

import time
from xlutils.copy import copy
import xlrd
import xlwt
import os
from playwright.sync_api import sync_playwright



# 初始化失败截图文件夹
os.system('rm -rf screenshot/*')
# 打开网页
browser = sync_playwright().start().chromium.launch(headless=False)
context = browser.new_context()
page = context.new_page()
# 读取excel文件里面的测试用例
workbook = xlrd.open_workbook('testcase.xls')
worksheet = workbook.sheet_by_index(0)
# 完成xlrd对象向xlwt对象转换，用于写入执行结果
excel = copy(wb=workbook)
excel_table = excel.get_sheet(0)
table = workbook.sheets()[0]

false_case = []

name = worksheet.row_values(0)
case_names = worksheet.col_values(0)[1:]
case_names = list(filter(None, case_names))
# print(case_names)
# 获取用例总行数
rows = worksheet.nrows
row = 1
# 根据用例名称为模块遍历用例步骤执行测试
for case_name in case_names:
    for i in range(row, rows):
        if worksheet.cell_value(i, 0) == '' or worksheet.cell_value(i, 0) == case_name:
            value = worksheet.row_values(i)
            case = dict(zip(name[1:], value[1:]))
            print(case['step'])
            test_values = list(case.values())[2:-3]
            test_values = list(filter(None, test_values))
            try:
                # 执行操作
                page.__getattribute__(case['method'])(*test_values)
                if case['assert_selector'] != '':
                    assert page.locator(case['assert_selector']).text_content() == case['assert_value']
                # 设置写入执行结果字体颜色
                style = xlwt.XFStyle()
                font = xlwt.Font()
                font.bold = True
                font.colour_index = 0x11
                style.font = font
                excel_table.write(i, 7, 'pass', style)
                # print(case['step'] + '----------\033[32mpass\033[0m')

            except:
                style = xlwt.XFStyle()
                font = xlwt.Font()
                font.bold = True
                font.colour_index = 0x0A
                style.font = font
                excel_table.write(i, 7, 'false', style)
                # 失败截图
                page.screenshot(path='screenshot/{}.png'.format(case['step']))
                # print(case['step'] + '----------\033[31mfalse\033[0m')
                false_case.append(case_name)
            time.sleep(1)
        else:
            row = i
            break
false_case = list(set(false_case))
case_num = len(case_names)
false_num = len(false_case)
print('总用例数----------', case_num)
print('成功----------{}----------{}%'.format(case_num-false_num, 100*int((case_num-false_num)/case_num)))
print('失败----------{}----------{}%'.format(false_num, 100-100*int((case_num-false_num)/case_num)))
# 写入执行结果保存文件
excel.save('testcase.xls')








# # 初始化失败截图文件夹
# os.system('rm -rf screenshot/*')
# # 打开网页
# browser = sync_playwright().start().chromium.launch(headless=False)
# context = browser.new_context()
# page = context.new_page()
# # 读取excel文件里面的测试用例
# workbook = xlrd.open_workbook('testcase.xls')
# worksheet = workbook.sheet_by_index(0)
# # 完成xlrd对象向xlwt对象转换，用于写入执行结果
# excel = copy(wb=workbook)
# excel_table = excel.get_sheet(0)
# table = workbook.sheets()[0]
# pass_num = 0
# false_num = 0
#
# name = worksheet.row_values(0)
# casename = worksheet.col_values(0)[1:]
# casename = list(filter(None, casename))
# print(casename)
# # 获取用例总行数
# rows = worksheet.nrows
# # 遍历用例步骤执行
# for i in range(1, rows):
#     if worksheet.cell_value(i, 0) == '' or casename[0]:
#         value = worksheet.row_values(i)
#         case = dict(zip(name[1:], value[1:]))
#         # print(case)
#         test_values = list(case.values())[2:-3]
#         test_values = list(filter(None, test_values))
#         try:
#             # 执行操作
#             page.__getattribute__(case['method'])(*test_values)
#             if case['assert_selector'] != '':
#                 assert page.locator(case['assert_selector']).text_content() == case['assert_value']
#             # 设置写入执行结果字体颜色
#             style = xlwt.XFStyle()
#             font = xlwt.Font()
#             font.bold = True
#             font.colour_index = 0x11
#             style.font = font
#             excel_table.write(i, 7, 'pass', style)
#             # print(case['step'] + '----------\033[32mpass\033[0m')
#             pass_num += 1
#         except:
#             style = xlwt.XFStyle()
#             font = xlwt.Font()
#             font.bold = True
#             font.colour_index = 0x0A
#             style.font = font
#             excel_table.write(i, 7, 'false', style)
#             # 失败截图
#             page.screenshot(path='screenshot/{}.png'.format(case['step']))
#             # print(case['step'] + '----------\033[31mfalse\033[0m')
#             false_num += 1
#         time.sleep(1)
#     else:
#         casename.pop(0)
#         break
# # 写入执行结果保存文件
# excel.save('testcase.xls')
# # print('总步骤数-------{}'.format(pass_num + false_num))
# # print('成功----{}----{}%'.format(pass_num, int(pass_num / (pass_num + false_num) * 100)))
# # print('失败----{}----{}%'.format(false_num, 100 - (int(pass_num / (pass_num + false_num) * 100))))
