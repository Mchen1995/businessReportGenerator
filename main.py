from funcLib import *

# 参数
version, head, filename, filename_of_demands = genArgs()

# 需求 读取 Excel
demand_sheet, number_of_demands = getDemandSheet(filename_of_demands)

# 创建文档及其开头的部分
document = createDoc(filename)

# 测试结果汇总表
createSummarizeTable(document, demand_sheet, number_of_demands)

document.add_heading('4. 测试项目内容：', 1)  # 一级标题：测试项目内容
document.add_heading('4.' + str(1) + ' 多元金融管理后台增加页签区配置', level=2)
# 测试案例表格
createCaseTable(document)

# 保存文档
document.save(filename)
