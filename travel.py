import docx
from docx import Document  # 导入库

path = 'C:/Users/sc/Desktop/test_docx/sample.docx'  # 文件路径
document = Document(path)  # 读入文件
tables = document.tables  # 获取文件中的表格集

for table in tables[:]:
    for i, row in enumerate(table.rows[:]):  # 读每行
        row_content = []
        for cell in row.cells[:]:  # 读一行中的所有单元格
            c = cell.text
            row_content.append(c)
        print(row_content)  # 以列表形式导出每一行数据
