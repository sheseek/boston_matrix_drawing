import os
import openpyxl
import pandas as pd
import matplotlib.pyplot as plt

def process_file(file_name):
    # 拼接完整的文件路径
    file_path = os.path.join("C:\\Users\\ThinkPad\\SynologyDrive\\Trademe\\", file_name + ".xlsx")

    # 打开Excel文件
    workbook = openpyxl.load_workbook(file_path)

    # 选择第一个工作表
    sheet = workbook.active

    # 获取数据
    data = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        data.append(row)

    # 关闭Excel文件
    workbook.close()

    # 将数据转换成DataFrame
    df = pd.DataFrame(data, columns=['SKU', '销量', '价格', '流量'])

    # 绘制气泡图
    scatter_plot = plt.scatter(df['流量'], df['销量'], s=df['价格']*10, alpha=0.5)

    # 在气泡上加上SKU的值
    for i, txt in enumerate(df['SKU']):
        plt.annotate(txt, (df['流量'][i], df['销量'][i]), textcoords="offset points", xytext=(0,5), ha='center')

    # 设置图表标题和轴标签
    plt.title('气泡图 - SKU位置')
    plt.xlabel('流量')
    plt.ylabel('销量')

    # 显示图例
    plt.legend(*scatter_plot.legend_elements(), title="价格")

    # 显示图表
    plt.show()

if __name__ == "__main__":
    # 用户输入文件名（不含后缀）
    user_input = input("请输入需要处理的文件名（不含后缀）: ")

    # 默认路径
    default_path = "C:\\Users\\ThinkPad\\SynologyDrive\\Trademe\\"

    # 补充后缀为".xlsx"
    file_name = user_input

    # 处理文件
    process_file(file_name)
