"""
生成第四章
"""
from funcLib import *

filename_of_cases, report_head, report_file, author, title_begin, case_begin = searchCase()
print("读取到文件：" + filename_of_cases)
print("正在生成：" + report_file)

# 创建文档
document = createDoc(report_head, False)

# 读取测试案例
case_sheet_list, _ = getCaseSheet(filename_of_cases)

for i in range(len(case_sheet_list)):
    sheet = case_sheet_list[i]
    createCaseParagraph(document, author, sheet, i + title_begin, i + case_begin)

# 保存文档
document.save(report_file)
print("完成")
