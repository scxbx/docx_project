import pandas as pd
import os

#文件路径
file_dir = r'C:\Users\sc\Desktop\to_merge'
#构建新的表格名称
new_filename = r'C:\Users\sc\Desktop\new.xlsx'
#找到文件路径下的所有表格名称，返回列表
file_list = os.listdir(file_dir)
new_list = []

for file in file_list:
    #重构文件路径
    file_path = os.path.join(file_dir,file)
    #将excel转换成DataFrame
    dataframe = pd.read_excel(file_path)
    dataframe.drop([len(dataframe) - 1], inplace=True)
    dataframe.drop(index=[0, 1], inplace=True)

    print(dataframe)
    #保存到新列表中
    new_list.append(dataframe)


#多个DataFrame合并为一个
df = pd.concat(new_list)
#写入到一个新excel表中
df.to_excel(new_filename,index=False)

