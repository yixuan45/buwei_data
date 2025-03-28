import pandas as pd
import numpy as np
from openpyxl import load_workbook

df_bumendaima = pd.read_excel('./整合数据.xlsx', sheet_name='部门代码')
df_shengji = pd.read_excel('./整合数据.xlsx', sheet_name='省级')
df_shiji = pd.read_excel('./整合数据.xlsx', sheet_name='市级')
df_quji = pd.read_excel('./区数据.xlsx')
df_quji['剩余'] = df_quji['剩余'].replace(np.nan, '')
df_quji['名称'] = df_quji['名称'] + df_quji['剩余']
df_quji = df_quji.fillna('')  # 已经把nan值转化为了空字符串
daima_list = []
name_list = []
quyu_list = []
zhixiashi = ['北京市', '天津市', '上海市', '重庆市']
shujuyuan_name = []  # 用于存储生成的数据源编码规则名字
shujuyuan_code = []  # 用于存储生成的数据源编码规则代码
qu_dict_li = []  # 用来存储区名整理成字典的列表
li = [2, 3, 4, 5, 6, 7, 8, 9, 'a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q',
      'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N',
      'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']
qu_dict = {}  # 将区级数据整理成字典的形式

for index_qu in range(2, 3560):  # 去遍历区数据
    if '市辖区' in df_quji.iloc[index_qu - 2, 1] or '省' in df_quji.iloc[index_qu - 2, 1]:  # 如果市辖区在这个单元格中
        continue
    if '市' in df_quji.iloc[index_qu - 2, 1] or '盟' in df_quji.iloc[index_qu - 2, 1] or '自治州' in df_quji.iloc[
        index_qu - 2, 1] or '安岭地区' in df_quji.iloc[
        index_qu - 2, 1]:  # 如果雄安新区,如果什么地区也算是一个区的话，就加一个'地区'in df_quji.iloc[index_qu-2,1]
        index_qu_old = index_qu  # 记录下市的名称
        qu_dict[df_quji.iloc[index_qu_old - 2, 1]] = []
        # 说明接下来的数据要存入列表中
        continue
    qu_dict[df_quji.iloc[index_qu_old - 2, 1]].append(df_quji.iloc[index_qu - 2, 1])
qu_dict['五家渠市'] = []
qu_dict['台湾省'] = []
qu_dict['香港特别行政区'] = []
qu_dict['澳门特别行政区'] = []
qu_dict['雄安新区'] = []
qu_dict['衡州市'] = []
qu_dict['满州里市'] = []
qu_dict['平潭综合实验区 '] = []
qu_dict['襄阳市'] = []
qu_dict['神农架区'] = []
qu_dict['湘西州'] = []
qu_dict['三沙市'] = []
qu_dict['阿坝州'] = []
qu_dict['甘孜州'] = []
qu_dict['凉山州'] = []
qu_dict['平坝市'] = []
qu_dict['盘州市'] = []
qu_dict['西秀市'] = []
# df_new = pd.DataFrame.from_dict(qu_dict, orient='index').T
# df_new.to_excel('./区域数据表.xlsx')
print(qu_dict)


def save_data_shengji(index_bumen, index_sheng):  # 用于存储省级表格的数据
    daima_list.append(df_bumendaima.iloc[index_bumen - 2, 0] + df_shengji.iloc[index_sheng - 2, 0].replace('B', ''))
    name_list.append(df_shengji.iloc[index_sheng - 2, 1] + df_bumendaima.iloc[index_bumen - 2, 1])


def save_data_shiji(index_bumen, index_shiji):  # 用于存储市级表格的数据
    daima_list.append(df_bumendaima.iloc[index_bumen - 2, 0] + df_shiji.iloc[index_shiji - 2, 0].replace('B', ''))
    name_list.append(df_shiji.iloc[index_shiji - 2, 1] + df_bumendaima.iloc[index_bumen - 2, 1])


def save_data_quji(index_bumen, index_quji, quji_code):  # 用于存储区级表格的数据
    daima_list.append(df_bumendaima.iloc[index_bumen - 2, 0] + quji_code)  # 要记得处理一下区级代码
    name_list.append(df_quji.iloc[index_quji - 2, 1] + df_bumendaima.iloc[index_bumen - 2, 1])


def china_data(index_sheng):
    for index_bumen in range(2, 42):
        save_data_shengji(index_bumen, index_sheng)
        quyu_list.append('部级')


for index_sheng in range(2, 42):  # 先遍历省级数据，再遍历部门代码
    sheng_name = df_shengji.iloc[index_sheng - 2, 1]
    for index_bumen in range(2, 62):
        if index_sheng == 2:  # 说明数据是中国
            china_data(index_sheng)
        # 先处理省级数据
        save_data_shengji(index_bumen, index_sheng)
        quyu_list.append('省级')
        if index_sheng > 34:  # 如果是长三角这种区域就不用去找市级了
            continue
    # 接下来处理市级数据表数据
    for index_shiji in range(2, 384):
        if '中国' in sheng_name:  # 如果中国在其中就不用再检查
            break
        if sheng_name not in df_shiji.iloc[index_shiji - 2, 1]:  # 如果省级名称没在市级当中，说明就不是省级对应的区
            continue
        for index_bumen in range(2, 62):
            if index_shiji > 369:  # 说明他没有区级记录了，像哪些兵工厂的名字
                daima_list.append(
                    df_bumendaima.iloc[index_bumen - 2, 0] + df_shiji.iloc[index_shiji - 2, 0].replace('B', ''))
                name_list.append(df_shiji.iloc[index_shiji - 2, 1] + df_bumendaima.iloc[index_bumen - 2, 1])
                quyu_list.append('市级')
                shujuyuan_code.append(df_shiji.iloc[index_shiji - 2, 0])
                shujuyuan_name.append(df_shiji.iloc[index_shiji - 2, 1])
                continue
            save_data_shiji(index_bumen, index_shiji)  # 如果不是以上的一种情况
            shujuyuan_code.append(df_shiji.iloc[index_shiji - 2, 0])
            shujuyuan_name.append(df_shiji.iloc[index_shiji - 2, 1])
            quyu_list.append('市级')
            # 如果没有执行if语句就说明，这个数据不是直辖市就要去找区级数据
            if sheng_name in zhixiashi:  # 如果是直辖市数据,就不用把sheng_name去掉
                judgment = df_shiji.iloc[index_shiji - 2, 1]
                qu_dict_li = qu_dict[judgment]  # 取出区对应的列表
                try:
                    qu_dict_li.remove('县')
                    qu_dict_li.remove('')
                except:
                    pass
                for data in qu_dict_li:
                    shiji_code = df_bumendaima.iloc[index_bumen - 2, 0] + str(
                        int(df_shiji.iloc[index_shiji - 2, 0].replace('B', '')) + 10)
                    shiji_name = df_shiji.iloc[index_shiji - 2, 1]
                    shujuyuan_name.append(shiji_name + data)  # 添加市级名称加区级名称
                    shiji_name = shiji_name + data + df_bumendaima.iloc[index_bumen - 2, 1]
                    shujuyuan_code.append('B' + str(int(df_shiji.iloc[index_shiji - 2, 0].replace('B', '')) + 10))
                    daima_list.append(shiji_code)
                    name_list.append(shiji_name)
                    quyu_list.append('区级')

            else:
                judgment = df_shiji.iloc[index_shiji - 2, 1].replace(sheng_name, '')  # 把市哪些给去掉，只剩下一个区名
                try:
                    qu_dict_li = qu_dict[judgment]  # 取出区对应的列表
                except:
                    if '市' in judgment or '盟' in judgment or '自治州' in judgment:
                        qu_dict[judgment] = []
                        qu_dict_li = qu_dict[judgment]
                try:
                    qu_dict_li.remove('县')
                    qu_dict_li.remove('')
                except:
                    pass
            qu_judgment = 0  # 用于去找li中的字母进行拼接
            for data in qu_dict_li:
                shiji_code = df_shiji.iloc[index_shiji - 2, 0]
                shiji_name = df_shiji.iloc[index_shiji - 2, 1]
                shiji_code = shiji_code[:len(shiji_code) - 1] + str(li[qu_judgment])  # 如果找到要的数据，就对代码进行替换，替换数据源编号的最后一位。
                shujuyuan_code.append(shiji_code)
                shujuyuan_name.append(shiji_name + data)
                shiji_name = shiji_name + data + df_bumendaima.iloc[index_bumen - 2, 1]
                shiji_code = df_bumendaima.iloc[index_bumen - 2, 0] + shiji_code.replace('B', '')  # 对编码进行一些处理
                daima_list.append(shiji_code)
                name_list.append(shiji_name)
                quyu_list.append('区级')

                qu_judgment += 1

data_dict = {
    '数据源编号': shujuyuan_code,
    '归属地': shujuyuan_name
}
df = pd.DataFrame(data_dict)
df.to_excel('数据源代码.xlsx')
print('存储完毕')

# data_dict = {
#     'daima': [],
#     'name': [],
#     'quyu': []
# }
#
# # Excel 文件路径
# file_path = '总数据.xlsx'
# for index_sheng in range(2, 42):
#     data_dict = {
#         'daima': [],
#         'name': [],
#         'quyu': []
#     }
#     sheng_name = df_shengji.iloc[index_sheng - 2, 1]  # 前面的名字
#     for index, name in enumerate(name_list):
#         if df_shengji.iloc[index_sheng - 2, 1] in name:
#             data_dict['daima'].append(daima_list[index])
#             data_dict['name'].append(name)
#             data_dict['quyu'].append(quyu_list[index])
#     if index_sheng > 2:
#         with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
#             df = pd.DataFrame(data_dict)
#             df.to_excel(writer, sheet_name=sheng_name, index=False)
#     else:
#         df = pd.DataFrame(data_dict)
#         df.to_excel('./总数据.xlsx', sheet_name=sheng_name)
#     print(f'{sheng_name}存储成功')
# print("执行完毕")
