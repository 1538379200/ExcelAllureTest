from BaseControl.base import BaseControl
from pathlib import Path
import os
import pandas as pd
import numpy as np
from ddt import ddt, data
import unittest
from HTMLTestRunner import HTMLTestRunner
"""
此代码将使用在unittest中
报告暂时使用为HTMLTestRunner，运行文件为根目录下的main.py文件
此代码仍以一个模块为一个用例加载
在main.py文件中运行后，报告文件将创建在根目录下unittest_report文件夹下，报告名为UnitResult.html
文件使用pandas作为数据处理模式，可以使用openpyxl代替，但是逻辑需要稍微整理重构，需保证数据形式相同
用例构建后的形式，一个模块分为一个list，形如：
[[模块1-用例1,模块2-用例2,...][模块2-用例1,模块2-用例2,...]]
构建此用例需要区别有模块有写入值和未写入值和不同值的区别，区分是否属于同一个模块
"""


keys = None
sys_name = None

excel_path = os.fspath(Path(__file__).resolve().parents[1]/'CaseData'/'NewExcelData.xlsx')
def get_cases():
    global sys_name
    # wb = openpyxl.load_workbook('NewExcelData.xlsx')
    sheet_and_value = pd.read_excel(io=excel_path, sheet_name=None)     # 获取excel所有sheet和value键值对
    sys_name = "默认系统名称"
    all_cases = []
    for sheet in list(sheet_and_value.keys()):
        head_row = sheet_and_value[sheet].iloc[0].index.values[0]       # 获取excel第一行的数据，如果不为未命名，则作为系统名称
        if head_row != 'Unnamed: 0':
            sys_name = head_row             # 获取系统名称
        sheet_value = sheet_and_value[sheet].values[1:]         # 将sheet、value键值对中，将value值取出来，过滤掉第二行title行
        # 将sheet名称放入列表中，作为后续可能增加的excel写入做准备
        value_insert_sheet = list(map(lambda x: list(np.insert(x, 0, sheet)), sheet_value))
        all_cases += value_insert_sheet     # 集合所有sheet用例
    all_cases = model_classify(all_cases)       # 将用例进行模块划分，划分函数在下
    return all_cases

def model_classify(case_list):
    all_cases = []
    filter_list = []
    for case in case_list:                      # 循环所有用例，用例形如[[用例1][用例2]]
        if not pd.isna(case[2]):                # 当用例的第三个值(excel中的模块列)不为空(nan)时
            if filter_list:
                all_cases.append(filter_list)   # 将数据添加进入all_cases列表
            filter_list = [case]                # 给filter_list赋值，使filter_list抛弃旧的数据重新进行其他模块用例收集
        else:
            filter_list.append(case)            # 当模块数据为空(nan)的时候，直接将用例添加到filter_list
    if filter_list:                             # 当循环最后一组时，最后一个数据大概率是空，所以需要单独进行判断并填入all_cases列表
        all_cases.append(filter_list)
    return all_cases

def run_cases(cases_list):
    global keys
    for case in cases_list:             # 循环所有用例，用例为模块分组后的用例
        case_params = case[4:8]         # 切片方式，取进行反射的数据参数
        filter_case_nan = [x for x in case_params if not pd.isna(x)]        # 反射数据参数剔除为空的数据
        if filter_case_nan[0] == "open_browser":        # 当第一个操作类型是open_browser时，为全局变量keys赋值BaseControl实例
            keys = BaseControl(filter_case_nan[-1])     # 最后一个输入参数，作为参数放入
        else:
            getattr(keys, 'event_'+filter_case_nan[0])(*filter_case_nan[1:])    # 将数据反射进入定义的基类方法中

@ddt
class TestExcel(unittest.TestCase):
    case_list = get_cases()         # 获取所有用例

    @data(*case_list)               # data进行用例放入
    def test_case(self, cases):
        run_cases(cases)            # 执行运行用例函数

def run():
    report_path = Path(__file__).resolve().parents[1] / 'unittest_report'
    report_path = os.fspath(report_path) + '/'
    discover = unittest.TestLoader().loadTestsFromTestCase(TestExcel)
    suite = unittest.TestSuite()
    suite.addTest(discover)
    with open(report_path+'UnitResult.html', 'wb') as fp:
        runner = HTMLTestRunner(stream=fp, title=sys_name, description=f"{sys_name}测试用例")
        runner.run(suite)
        fp.close()

# if __name__ == '__main__':
#     run()


