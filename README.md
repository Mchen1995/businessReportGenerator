# businessReportGenerator
本 Python 脚本程序用于生成集团门户软件开发业务测试报告。

使用前提：
1. 需要读取测试案例和软件下发需求两个 Excel 文件，分别命名为 "cases.xls" 和 "demands.xls"。
2. 测试案例由多个 sheet 组成，每个 sheet 以测试人员的姓名命名，以便程序自动填入测试报告。每个 sheet 由 A, B, C 三列组成，分别对应业务测试报告中的 4.x, 4.x.x, 4.x.x.x 级别。注：在两个子需求（4.x.x 级别）之间，C 列应空一行。

使用方法：
1. IDE 虚拟环境运行：运行 main.py，即可生成业务测试报告
2. 双击 main.exe 即可

###### 生成的报告名为 report.docx，然后手动插入截图即可。