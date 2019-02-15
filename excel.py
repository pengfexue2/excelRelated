import xlrd
import xlsxwriter
import os

#要整理的excel表格 xls、xlsx文件所在文件夹
xpath = "C:\\Users\\TED\\Desktop\\excel\\excel文件夹"

typedata = []
name = []
#用来读取列识别列表中xls或xlsx文件，将其名字添加到list中返回
def collect_xls(list_collect):
    for each_element in list_collect:
        if isinstance(each_element,list):
            collect_xls(each_element)
        elif each_element.endswith("xls"):
            typedata.insert(0,each_element)
        elif each_element.endswith("xlsx"):
            typedata.insert(0,each_element)
    return typedata
#读取文件夹中包含的所有xls和xlsx格式表格文件
def read_xls(path):
    for file in os.walk(path):
        # os.walk() 返回三个参数：路径，子文件夹，路径下的文件
        for each_list in file[2]:
            file_path = file[0]+"/"+each_list
            name.insert(0,file_path)
        all_xls = collect_xls(name)

    return all_xls

src = read_xls(xpath)
print(src)
total = [['部门名称', '招考职位', '职位简介', '招考人数', '专业', '学历','来源']]

for xls_item in src:
    wb = xlrd.open_workbook(xls_item)
    sheets = wb.sheet_names()

    for index in range(len(sheets)):
        table = wb.sheets()[index]
        nrows = table.nrows
        if nrows == 0:
            continue

        if xls_item =="C:\\Users\\TED\\Desktop\\excel\\excel文件夹/2019.xls":
            label = table.row_values(1)
            item1 = label.index("部门名称")
            item2 = label.index("招考职位")
            item3 = label.index("职位简介")
            item4 = label.index("招考人数")
            item5 = label.index("专业")
            item6 = label.index("学历")

            for i in range(2, nrows):
                item = [table.row_values(i)[item1], table.row_values(i)[item2], table.row_values(i)[item3],
                        table.row_values(i)[item4], table.row_values(i)[item5], table.row_values(i)[item6]]
                item.append(xls_item)
                total.append(item)
        else:
            label = table.row_values(0)

            item1 = label.index("招录机关")
            item2 = label.index("招考职位")
            item3 = label.index("职位简介")
            item4 = label.index("招考人数")
            item5 = label.index("专业")
            item6 = label.index("学历")

            for i in range(1,nrows):
                item=[table.row_values(i)[item1],table.row_values(i)[item2],table.row_values(i)[item3],table.row_values(i)[item4],table.row_values(i)[item5],table.row_values(i)[item6]]
                item.append(xls_item)
                total.append(item)

workbook = xlsxwriter.Workbook("result.xlsx")
worksheet = workbook.add_worksheet()

group=['A','B','C','D','E','F','G']
for i in range(len(total)):
    for j in range(len(total[i])):
        worksheet.write(f"{group[j]}{i+1}",total[i][j])
workbook.close()



