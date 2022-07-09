from docx.shared import Inches
from docx import Document
import xlrd


# 读取需求 Excel
def getDemandSheet(filename_of_demands):
    demand_sheet = xlrd.open_workbook(filename_of_demands).sheet_by_index(0)
    count = demand_sheet.nrows - 1  # 需求个数
    return demand_sheet, count


# 创建文档
def createDoc(filename):
    document = Document()
    document.add_heading(filename, 0)  # 标题
    document.add_heading('1. 测试日期：', 1)  # 一级标题：测试日期
    document.add_heading('2. 测试人员：黄乃芳、郑杰、纪雅容、贺东琴', 1)  # 一级标题：测试人员
    document.add_heading('3. 测试结果：符合需求，测试通过', 1)  # 一级标题：测试结果
    return document


# 生成参数
def genArgs():
    version = "2.5.7.11"  # 版本号
    head = '集团门户业务测试报告_' + version  # 全文标题
    filename = head + '.docx'  # 测试报告文件名
    filename_of_demands = 'demand.xls'  # 需求 Excel
    return version, head, filename, filename_of_demands


# 生成测试结果汇总表格
def createSummarizeTable(document, demand_sheet, count):
    table = document.add_table(rows=1, cols=4, style='Table Grid')
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = '序号'
    hdr_cells[1].text = '需求编号'
    hdr_cells[2].text = '测试需求点'
    hdr_cells[3].text = '测试结果'
    for i in range(count):
        row_cells = table.add_row().cells
        row_cells[0].text = str(i + 1)  # 序号
        row_cells[1].text = demand_sheet.cell_value(i + 1, 1)  # 需求编号
        row_cells[2].text = demand_sheet.cell_value(i + 1, 2)  # 测试需求点
        row_cells[3].text = '通过'  # 序号
    table.style = 'Colorful List'
    for cell in table.columns[0].cells:
        cell.width = Inches(0.5)
    for cell in table.columns[1].cells:
        cell.width = Inches(3)
    for cell in table.columns[2].cells:
        cell.width = Inches(4)
    for cell in table.columns[3].cells:
        cell.width = Inches(1.5)


# 生成测试案例表格
def createCaseTable(document):
    table = document.add_table(rows=1, cols=4, style='Table Grid')
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = '测试人'
    hdr_cells[0].style = 'red'
    hdr_cells[1].text = ''
    hdr_cells[2].text = '编写人'
    hdr_cells[3].text = ''

    row_cells = table.add_row().cells
    row_cells[0].text = '测试用例编号'
    row_cells[1].text = ''
    row_cells[2].text = '测试日期'
    row_cells[3].text = ''

    row_cells = table.add_row().cells
    row_cells[0].text = '测试用例名称'
    row_cells[1].text = ''
    table.cell(2, 1).merge(table.cell(2, 2)).merge(table.cell(2, 3))  # 测试用例名称行合并为 2 列

    row_cells = table.add_row().cells
    row_cells[0].text = '测试目标'
    row_cells[1].text = '功能正常'
    table.cell(3, 1).merge(table.cell(3, 2)).merge(table.cell(3, 3))  # 测试目标行合并为 2 列

    row_cells = table.add_row().cells
    row_cells[0].text = '预期结果：功能正常'
    table.cell(4, 0).merge(table.cell(4, 1)).merge(table.cell(4, 2)).merge(table.cell(4, 3))  # 预期结果行合并为 1 列

    row_cells = table.add_row().cells
    row_cells[0].text = '实际结果：【此处应截屏说明】\n'
    table.cell(5, 0).merge(table.cell(5, 1)).merge(table.cell(5, 2)).merge(table.cell(5, 3))  # 实际结果行合并为 1 列

    row_cells = table.add_row().cells
    row_cells[0].text = '结果及意见'
    row_cells[1].text = '测试结果符合需求，测试通过。'
    table.cell(6, 1).merge(table.cell(6, 2)).merge(table.cell(6, 3))  # 结果及意见行合并为 2 列

    row_cells = table.add_row().cells
    row_cells[0].text = '备注'
    row_cells[1].text = ''
    table.cell(7, 1).merge(table.cell(7, 2)).merge(table.cell(7, 3))  # 备注行合并为 2 列
