import openpyxl
import allure
from BaseControl.base import BaseControl
import pytest
from pathlib import Path
import os

keys = None
old_story = None

def get_cases():
    wb = openpyxl.load_workbook('NewExcelData.xlsx')
    sheets_name = wb.sheetnames
    sys_name = "默认系统名称"
    all_cases = []
    filter_list = []
    for sheet in sheets_name:
        ws = wb[sheet]
        if list(ws.values)[0][0] is not None:
            sys_name = list(ws.values)[0][0]    # 重新赋值的系统名称
        no_title_list = list(ws.values)[2:]     # 过滤title的用例信息
        all_cases = all_cases + no_title_list
    case_list = []
    for case in all_cases:                      # 将用例按照模块编写划分，划分为多个列表嵌套形式
        if case[1] is not None:
            if case_list:
                filter_list.append(case_list)
            case_list = [case]
        else:
            case_list.append(case)
    if case_list:
        filter_list.append(case_list)
    return sys_name, filter_list


def run_cases(case_list):
    """
    在此文件中，将确实所谓的模块概念，如需要加入模块概念，则需要在excel中继续添加一个‘模块’列，原模块列作为用例名存在
    此文件中的模块名和allure的feature属于相同概念，可以不添加feature，只使用用例显示
    如果想在添加系统模块的列，则在列表中添加后，需要再次进行数据处理提取
    :param case_list: 运行的用例列表
    :return:
    """
    global keys, old_story
    old_title = None                                            # 旧的用例标题，在未编写用例名称时使用
    for case in case_list:
        if case[1] is not None:
            old_story = case[1]                                 # 设置旧的story为excel的模块名
        # allure.dynamic.story(old_story)
        case_params = [x for x in case[3:7] if x is not None]   # 取出进行运行的用例参数
        print(case_params)
        if case[1] is not None:                                 # 当模块名不为空，将excel的模块名作为用例名
            old_title = case[1]
        allure.dynamic.title(old_title)
        with allure.step(case[-1]):
            if 'open_browser' in case_params:
                keys = BaseControl(case_params[-1])
            else:
                getattr(keys, "event_"+case_params[0])(*case_params[1:])

class TestExcel:
    sys_name, cases = get_cases()

    @allure.epic(sys_name)                      # 设置运行的系统名称
    @pytest.mark.parametrize("case", cases)
    def test_cases(self, case):
        # allure.dynamic.feature(case[0][1])      # 设置模块名称，在此文件中模块与用例为相同概念
        run_cases(case)


if __name__ == '__main__':
    report_path = Path(__file__).resolve().parents[1] / 'reports' / 'results'
    pytest.main(["-s", "-v", "read_excel_step.py", "--alluredir", "../reports/results", '--clean-alluredir'])
    os.system(f'allure serve {report_path}')
