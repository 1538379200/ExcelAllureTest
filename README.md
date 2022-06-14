# ExcelAllureTest
excel数据驱动结合allure生成报告代码

## 代码说明
目前代码并不是最终版本，但可以在此基础上进行优化和删减，满足个人需求
### read_excel_pandas.py
此文件主要使用pandas进行数据处理，具有集合allure生成报告的功能

### read_excel_openyxl.py
此文件主要使用openpyxl进行excel数据处理，在pandas的基础上，也增加了写入excel的功能，当excel数据中存在assert关键字，将进行断言，断言失败会将用例在excel中标红，同时allure报告中也会raise成失败用例，运行结果的excel不会覆盖原有用例数据，而是在reports目录下生成一个TestResult.xlsx文件

### read_excel_step.py
此文件是openpyxl进行数据处理，但是并未增加写入功能，与前面逻辑稍有不同，前面的文件可能会不符合allure的用例显示设计逻辑，而此文件则是为了更符合allure的逻辑，但是excel中的数据会出现变化，此文件中的模块相当于一个用例，如果想再增加模块，则自行在excel中添加一个模块列，将原模块列设置成用例名称，数据提取方式需要稍作修改

### read_excel_unittest.py
此文件作为unittest结合HTMLTestRunner生成报告，仅做写入excel准备，未完成excel写入功能
