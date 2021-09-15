## docx_project

#### household_residence_census.household_residence_census_generator.py

将**汇总表**转为**农村宅基地户籍情况调查表**，用法如下：

1. 将**汇总表**放入**summary**文件夹中（可放入文件夹，但必须保证里面的文件都是**汇总表**）。

2. 关闭所有excel

3. 打开**household_residence_census_generator**。

4. 在**result**文件夹中获取**农村宅基地户籍情况调查表**。



   #### registration_form_for_share_certificate_application.registration_share.py

将**汇总表**转为**股权证申领登记表**，用法如下：

1. 将**汇总表**放入**summary**文件夹中。

2. 关闭所有excel

3. 打开**registration_share.py**。

4. 在**result**文件夹中获取**股权证申领登记表**。



#### menu_generator.py

将**汇总表**转为**目录**，用法如下：

1. 确保**A2**的形式为**集体经济组织名称：xx县xx镇xx村xx组股份经济合作社**
2. 将**汇总表**放入**summary**文件夹中。
3. 关闭所有excel。
4. 打开**menu_generator.py**。
5. 在**menu**文件夹中获取**目录**。



#### menu_all_in_one.py

对**menu_generator.py**生成的**目录**的格式进行调整，以减少页数。用法如下：

1. 将**menu_generator.py**生成的**目录**放入**old menu**文件夹。
2. 关闭所有excel。
3. 打开**menu_all_in_one.py**。
4. 在 **new menu**文件夹中获取**新目录**。



#### pandas_merge.py

将多个汇总表合并成一个大汇总表，并输出重复的身份证号码。用法如下：

1. 在桌面新建**to_merge**文件夹。
2. 将要合并的汇总表放入**to_merge**文件夹中。
3. 关闭所有excel。
4. 打开**pandas_merge.py**，输入数据开始的行数。
5. 在桌面获取**大汇总表.xlsx**和**重复身份证.txt**。



#### produce_confirm.py

将**确认汇总表**转为**确认表**。用法如下：

1. 将**确认汇总表**放入**summary**文件夹中。
2. 关闭所有excel。
3. 打开**produce_confirm.py**。
4. 在**confirm**文件夹中获取**确认表**。
5. 如果生成失败，请检查确认汇总表内是否有空值（联系电话和备注可以为空）。



#### share_warrant_generator.py

<img src=".\pictures\share_warrant_generator.PNG" alt="share_warrant_generator"  />

结合用户在图形界面输入的信息，将确认表生成**股权证**。用法如下：

1. 点击**路径选择**，并选取需要的**确认表**。
2. 填写剩下的所有信息。
3. 点击**生成股权证。**
4. 在**warrant**文件夹中获取股权证。