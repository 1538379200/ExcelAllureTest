import pandas as pd
import pytest
from BaseControl.base import BaseControl
import allure
from pathlib import Path
import os
"""
代码说明：
1. 表格用例的层级划分：系统->模块->流程->用例(allure表示: epic->feature->story->step)
2. 表格数据相对固定，如若增加优先级等数据，则在用例parametrize中需要进行添加
3. 此代码基本使用allure报告形式管理用例结果，可增加excel写入方式
4. 代码认定多sheet也为一个系统的用例，所以即使写了多个sheet，仍会取出所有值，并进行顺序执行
5. 原代码excel使用pandas进行数据处理，openyxl可查看read_excel_openyxl.py
"""

keys = None                 # 全局变量key，用来接受BaseControl实例
old_model = None            # 全局变量old_model，用来保存旧数据的模块名称
old_story = None            # 全局变量old_story，用来保存旧的story(流程名称)

def get_cases():
    df = pd.read_excel(io='NewExcelData.xlsx', sheet_name=None)   # 读取excel表格，获取sheet和数据
    all_cases = []
    sys_name = "未获取到的默认值"
    for sheet in list(df.keys()):                           # 循环所有sheet
        sys_name = df[sheet].head().iloc[0].index[0]        # 获取表格系统名称
        sheet_data = df[sheet].iloc[:, 1:]                  # 通过键值对，取出sheet中所有值
        all_cases = all_cases + list(sheet_data.values)     # 将取出的值添加进入all_cases列表中
    case_data = [list(x) for x in all_cases][1:]            # 分解列表，并转换为python的list类型，截取第二行数据
    return sys_name, case_data

class TestExcel:
    sys_name, data_list = get_cases()               # 获取取出的表格数据

    @allure.epic(sys_name)
    @pytest.mark.parametrize('model, process, event, location, elements, txt, desc', data_list)
    def test_cases(self, model, process, event, location, elements, txt, desc):
        global keys, old_story, old_model           # 指定全局变量
        if pd.isna(model):                          # 当获取的数据为None时，使用旧数据作为模块名
            case_model = old_model
        else:
            case_model = model
        allure.dynamic.feature(case_model)          # 动态设置allure的模块名称
        if pd.isna(process):                        # 当流程为None时，设置story为旧数据
            case_story = old_story
        else:
            case_story = process
        allure.dynamic.story(case_story)            # 动态设置allure的story
        allure.dynamic.title(case_story)            # 动态设置用例的标题
        action_data = [x for x in [location, elements, txt] if not pd.isna(x)]  # 过滤为None的空用例数据
        if event == "open_browser":
            keys = BaseControl(txt)
        else:
            with allure.step(desc):                 # 设置用例的step，以表格的描述为准
                getattr(keys, 'event_' + event)(*action_data)
        old_story = case_story                      # 设置旧的story
        old_model = case_model                      # 设置旧的model

if __name__ == '__main__':
    report_path = Path(__file__).resolve().parents[1]/'reports'/'results'
    pytest.main(['-s', '-v', 'read_excel_pandas.py', '--alluredir', '../reports/results', '--clean-alluredir'])
    os.system(f'allure serve {report_path}')