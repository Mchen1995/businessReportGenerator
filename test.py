from funcLib import *

filename_of_cases = 'cases.xls'
sheet_list, name_list = getCaseSheet(filename_of_cases)
column = sheet_list[0].col_values(2, 0)
index = get_min_titles_index(column, 0)
print(column)
print(index)
