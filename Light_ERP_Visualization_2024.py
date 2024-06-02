# -*- coding: utf-8 -*-

#######################################
#                                     #
#  ERP比赛 操作时间线 可视化图像生成  #
#  适用于 数智企业经营管理沙盘        #
#                                     #
#######################################

import random
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.gridspec as gridspec

# Excel文件路径
# 智能分析 - 财务报表 - 现金流量表 - 实时，查询，复制内容到 excel 表
# 不同小组分别粘贴到不同工作表，保持工作表名称 "Sheet1"、"Sheet2" 不变，方便替换 "Sheet" 为 "小组"
# 各工作表 第一行 列名:  编号	动作	资金	剩余	时间	岗位	备注

excel_file_path = 'savedExcel.xlsx'

# 输出图像前缀日期
dateStr = "0529_"

# 输出图像size
figWidth = 140 # 根据操作数量调节
figHeight = 10
myDpi = 90 # DPI (dots per inch)

# 添加数据标签，此例为备注
addDataLabel_Boolean = True

# 读取Excel文件并获取所有工作表的名称
xls = pd.ExcelFile(excel_file_path)
sheet_names = xls.sheet_names

# 创建一个字典，用于存储每个工作表的DataFrame
dfs = {}

# 遍历工作表名称，为每个工作表创建一个DataFrame
for sheet_name in sheet_names:
    dfs[sheet_name] = pd.read_excel(excel_file_path, sheet_name=sheet_name)

# 创建颜色映射字典，用于根据岗位改变柱状图颜色
position_color_map = {'人力资源总监': 'blue', '项目总监': 'green', '营销总监': 'red', '运营总监': 'orange', '总经理': 'brown'}

# 定义一个函数来添加x轴分段填充颜色，根据x轴标签是否相同，循环改变
def add_segment_colors_Toggle(ax, df):
    # 初始化一个新变量，用于在0和1之间切换
    toggle_value = 0
    segment_color = 'lightgray'

    # 遍历 data 列表，检查相邻元素是否变化
    dataTime = df['时间']
    for i, time in enumerate(dataTime):
        if i > 0:  # 确保这不是第一个元素
            # 检查当前值是否与前一个值不同
            if time != dataTime[i-1]:
                # 如果不同，则切换 toggle_value
                toggle_value = 1 - toggle_value
            if toggle_value == 0:
                segment_color = 'lightgray'
            else:
                segment_color = 'white'                
            rect = mpatches.Rectangle((i - 0.5, 0), 1, ax.get_ylim()[1], linewidth=0, color=segment_color, alpha=0.2)
            ax.add_patch(rect)                 
        else:  # 第一个元素            
            # 创建矩形，宽度为1，中心与柱状图中心对齐
            rect = mpatches.Rectangle((i - 0.5, 0), 1, ax.get_ylim()[1], linewidth=0, color=segment_color, alpha=0.2)
            ax.add_patch(rect)  # 添加到轴上

# 定义一个函数来绘制柱状图
def plot_bars(ax, df, column, position_color_map):
    for i, (time, value, position, note) in enumerate(zip(df['时间'], df[column], df['岗位'], df['备注'])):
        color = position_color_map[position]  # 获取岗位对应的颜色
        ax.bar(i, value, color=color, alpha=0.7)  # 绘制柱状图        
        # 添加备注作为数据标签
        if addDataLabel_Boolean:
            if note:  # 如果备注不为空
                ax.text(i, value + 0.05 * df[column].max(), note, ha='center', va='bottom', fontweight='bold')

    # 添加图例
    for position, color in position_color_map.items():
        ax.bar(0, 0, color=color, label=position)  # 添加岗位图例
    ax.legend(facecolor='white', loc='upper right', fontsize=9)  # 显示图例，背景颜色为灰色，位置在图表右上角

# 如果相邻的x轴标签前8位相同，则合并标签并居中
def merge_similar_xticks(ax, df, tick_spacing=9):
    # 获取时间列并转换为字符串
    time_strings = df['时间'].astype(str)
    
    # 初始化上一个标签的前8位和当前标签的索引
    prev_tick_prefix = None
    tick_start_index = 0
    merged_ticks = []  # 存储合并后的标签和位置
    
    # 遍历每个标签
    for i, time in enumerate(time_strings):
        current_tick_prefix = time[:tick_spacing]
        
        # 如果当前标签的前8位与上一个不同，或者这是第一个标签
        if current_tick_prefix != prev_tick_prefix or prev_tick_prefix is None:
            # 如果这不是第一个标签，就处理前一个合并的标签
            if prev_tick_prefix is not None:
                # 计算标签的中心位置
                center_index = (tick_start_index + i - 1) / 2
                # 存储合并的标签和位置
                merged_ticks.append((center_index, prev_tick_prefix))
            
            # 更新当前标签的前8位和开始索引
            prev_tick_prefix = current_tick_prefix
            tick_start_index = i
            
    # 处理最后一个标签
    if prev_tick_prefix is not None:
        center_index = (tick_start_index + len(time_strings) - 1) / 2
        merged_ticks.append((center_index, prev_tick_prefix))

    # 设置新的x轴刻度和标签
    ax.set_xticks([x[0] for x in merged_ticks])
    ax.set_xticklabels([x[1] for x in merged_ticks], rotation=0, ha='right')
  
# 所有小组 分别输出图 

for key, value in dfs.items():
    fig_title_text = key.replace('Sheet', '小组')
    df = value
    
    # 上下排列（nrows=2, ncols=1）
    fig, axes = plt.subplots(nrows=2, ncols=1, figsize=(figWidth, figHeight))
    # 添加总标题
    fig.suptitle(fig_title_text + ' 资金和剩余随时间变化图', fontsize=12, fontweight='bold')

    # 绘制第一个柱状图：资金
    plot_bars(axes[0], df, '资金', position_color_map)
    add_segment_colors_Toggle(axes[0], df)
    merge_similar_xticks(axes[0], df)  # 调用函数合并标签
    axes[0].set_title('资金随时间变化\n', fontsize=9, loc='left') 

    # 取消第一个图表的边框，但保留x轴和y轴
    for spine in ['top', 'right']:  # 只隐藏上边框和右边框
        axes[0].spines[spine].set_visible(False)

    # 绘制第二个柱状图：剩余
    plot_bars(axes[1], df, '剩余', position_color_map)
    add_segment_colors_Toggle(axes[1], df)
    merge_similar_xticks(axes[1], df) 
    axes[1].set_title('剩余随时间变化\n', fontsize=9, loc='left')

    # 取消第二个图表的边框，但保留x轴和y轴
    for spine in ['top', 'right']: 
        axes[1].spines[spine].set_visible(False)
    
    # Save
    plt.savefig(dateStr + fig_title_text + '.png', dpi=myDpi, bbox_inches='tight', pad_inches=0.1) 

    # 显示图表
    plt.show()

print('\n-------- End --------\n')
