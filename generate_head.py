"""
仅生成前三章
"""
from funcLib import *

# 参数
head, filename, filename_of_demands, filename_of_cases = genArgs()
print("正在生成：" + filename)

# 读取需求 Excel
demand_sheet, number_of_demands = getDemandSheet(filename_of_demands)
print("已读取：" + filename_of_demands)

# 创建文档及其开头的部分
document = createDoc(head, True)

# 测试结果汇总表
createSummarizeTable(document, demand_sheet, number_of_demands)

document.add_heading('4. 测试项目内容：', 1)  # 一级标题：测试项目内容

# 保存文档
document.save(filename)
print("完成")
