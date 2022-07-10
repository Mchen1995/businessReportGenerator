from funcLib import *

# 获取键盘输入版本号
version = input("输入版本号: ")
print("正在生成：集团门户业务测试报告_" + version + ".doc")

# 参数
head, filename, filename_of_demands, filename_of_cases = genArgs(version)

# 读取需求 Excel
demand_sheet, number_of_demands = getDemandSheet(filename_of_demands)
print("已读取：" + filename_of_demands)

# 读取测试案例 Excel
case_sheet_list, name_list = getCaseSheet(filename_of_cases)
print("已读取：" + filename_of_cases)

# 创建文档及其开头的部分
document = createDoc(head)

# 测试结果汇总表
createSummarizeTable(document, demand_sheet, number_of_demands)

document.add_heading('4. 测试项目内容：', 1)  # 一级标题：测试项目内容

# 测试案例段落
for i in range(len(case_sheet_list)):
    sheet = case_sheet_list[i]
    author = name_list[i]
    createCaseParagraph(document, author, sheet, i + 1)

# 保存文档
document.save(filename)
print("完成")
