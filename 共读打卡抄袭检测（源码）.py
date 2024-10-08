# 导入所需的库
import pandas as pd
from fuzzywuzzy import fuzz
import time
import os
import pathlib
import tkinter as tk
import time
from threading import Thread
current_directory = pathlib.Path().absolute()
current_directory =str(current_directory)+"./data/"
print("当前工作目录：", current_directory)
file_names = os.listdir(current_directory)

filename=""
for file in file_names:
    if file.endswith('.xlsx'):
        print(file)
        filename=file

# similarity = fuzz.ratio("小明玩耍的时候，不小心把花瓶打碎了", "花瓶被小明在玩耍的时候不小心打碎了")
# print(similarity)
similarity=1
now = time.localtime()
print(time.strftime("%m-%d %H:%M", now))
# 读取 Excel 文件
df=pd.read_excel("./data/"+filename)

# 获取“简介”列的数据
intro = df["简介"][0:]
print(intro.head())
# 创建一个空列表来存储结果
result = []
root = tk.Tk()
root.title("处理进度")

label = tk.Label(root, text=f"共"+str(len(intro))+"条数据，当前处理进度为0 / "+str(len(intro)))
label.pack()

def update_label():
# 遍历“简介”列的每一个单元格
    for i in range(len(intro)):
        label.config(text=f"共"+str(len(intro))+"条数据，当前处理进度为"+str(i+1)+"/"+str(len(intro)))
        root.update()

        # 获取当前单元格的文本数据
        # print("单元格:",i)
        current = intro[i]
        # 定义查重对比范围
        start = max(0, i)
        end = min(len(intro), i + 9500)
        # 遍历对比范围内的其他单元格
        for j in range(start, end):
            # 跳过当前单元格
            if j == i:
                continue
            # 获取其他单元格的文本数据
            other = intro[j]
            # 计算两个文本数据的相似度
            similarity = fuzz.ratio(current, other)
            # 如果相似度大于等于50%，则将当前单元格的索引序号和相似单元格的索引序号添加到结果列表中
            if df.iloc[i][14] == "审核不通过" or df.iloc[j][14] == "审核不通过":
                continue
            if similarity >= 70:
                result.append(((i, j),similarity))

    def close_window():
        root.destroy()


    close_button = tk.Button(root, text="关闭", command=close_window)
    close_button.pack()

update_label()

root.mainloop()
# 创建一个ExcelWriter对象，并指定文件名和模式（w表示写入，a表示追加）
now = time.localtime()
# 使用 ExcelWriter 保存 DataFrame 到 Excel 文件
with pd.ExcelWriter("暑假共读打卡抄袭证据表_截止"+time.strftime("%m-%d %H", now)+".xlsx", engine='openpyxl') as writer:
    # df.to_excel(writer, sheet_name='Sheet1')
# writer = pd.ExcelWriter("寒假共读打卡抄袭证据表_截止"+time.strftime("%m-%d %H", now)+".xlsx", mode="w")
    data=pd.DataFrame(result,columns=["编号组","相似百分比(%)"])
    df.to_excel(writer,sheet_name="打卡数据"+time.strftime("%m-%d %H", now))
    data.to_excel(writer,sheet_name="疑似抄袭编号组别")

    i=0
    count_dict = {}
    for part in result:
        index_tuple=part[0]
        sim=part[1]
        row1 = df.iloc[index_tuple[0]]
        row2 = df.iloc[index_tuple[1]]
        # 将两行数据合并为一个新的数据框
        new_df = pd.concat([row1, row2], axis=1)  # concat方法用于合并数据，axis=1表示按列合并
        new_df.to_excel(writer,sheet_name=str(i)+"_"+str(sim)+"%")
        i=i+1

        name= "_姓名"+str(df.iloc[index_tuple[0]][3])+"_昵称"+str(df.iloc[index_tuple[0]][2])+"_手机号"+str(df.iloc[index_tuple[0]][4])+"_院系"+str(df.iloc[index_tuple[0]][18])+"_专业"+str(df.iloc[index_tuple[0]][19])+"_学号"+str(df.iloc[index_tuple[0]][20])
        # print(name)
        if name in count_dict:
            # 如果字符已经在字典中，就将其对应的值加一
            count_dict[name] += 1
            # 如果字符不在字典中，就将其添加到字典中，并将其对应的值设为一
        else:
            count_dict[name] = 1

        # 输出字典中的每个键值对，即每个字符及其出现的次数
    # print(count_dict)
    count_dictzip = zip(count_dict.keys(), count_dict.values())
    cout=pd.DataFrame(count_dictzip,columns=["信息","抄袭次数"])
    # 使用 ExcelWriter 保存 DataFrame 到 Excel 文件
    # with pd.ExcelWriter('output.xlsx', engine='openpyxl') as writer:
    #     df.to_excel(writer, sheet_name='Sheet1')
    cout.to_excel(writer,sheet_name="抄袭汇总统计")
    # writer.save()
# print(time.strftime("%m-%d %H:%M", now))




