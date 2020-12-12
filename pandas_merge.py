import winreg
from collections import Counter

import pandas as pd
import os


def findDuplicatedElements(mylist):
    b = dict(Counter(mylist))
    return [key for key, value in b.items() if value > 1]  # 只展示重复元素
    # print({key: value for key, value in b.items() if value > 1})  # 展现重复元素和重复次数


def get_desktop():
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                         r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders')
    return winreg.QueryValueEx(key, "Desktop")[0]


row_to_start = int(input('请输入数据开始的行数：'))
drop_index_list = []
# 汇总表从第 n 行开始有数据 出去表头还要drop n-2 行
for i in range(0, row_to_start - 2):
    drop_index_list.append(i)


# 文件路径
file_dir = os.path.join(get_desktop(), 'to_merge')

# 构建新的表格名称
new_filename = os.path.join(get_desktop(), '大汇总.xlsx')
# 找到文件路径下的所有表格名称，返回列表
file_list = os.listdir(file_dir)
new_list = []

for file in file_list:
    # 重构文件路径
    file_path = os.path.join(file_dir, file)
    # 将excel转换成DataFrame
    dataframe = pd.read_excel(file_path)
    dataframe.drop([len(dataframe) - 1], inplace=True)
    dataframe.drop(index=drop_index_list, inplace=True)

    # print(dataframe)

    # 保存到新列表中
    new_list.append(dataframe)

# 多个DataFrame合并为一个
df = pd.concat(new_list)
# 写入到一个新excel表中
df.to_excel(new_filename, index=False)
print('总计：', len(df), '户')

# for i in range(len(df)):
# print(df.iloc[0: len(df), 6])
my_list = df.iloc[0: len(df), 6].tolist()
# print(my_list)
duplicatedList = findDuplicatedElements(my_list)
print('以下身份证重复：')

f = open(os.path.join(get_desktop(), '重复身份证.txt'), 'w')
for duplicatedItem in duplicatedList:
    print(duplicatedItem)
    f.write(str(duplicatedItem))
    f.write('\n')
f.close()

input('Press enter to quit program.')
