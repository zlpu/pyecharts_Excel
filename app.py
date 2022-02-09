# coding:utf-8
"""
1.题目：基于python的pyecharts模块实现将excel中的数据可视化
2.web框架：Falsk
数据刷新方式：excel中的数据发生改变后需要重新运行此app.py文件
3.用到的包：pyecharts、xlrd、flask、datatime
4.步骤：
------01：导入相关的包
------02：数据源的准备及处理
------03：作图,图表的相关属性设置
------04.图表组合，Grid
------05.flask（app.route）
------06.run app.py
Creator: pzl960504@163.com
"""
'''导入相关的包'''
from flask import Flask
from jinja2 import Markup, Environment, FileSystemLoader
from pyecharts.globals import CurrentConfig

CurrentConfig.GLOBAL_ENV = Environment(loader=FileSystemLoader("./templates"))
from pyecharts.charts import Map3D, Map, Bar, Pie, WordCloud, \
    Line, Gauge, Geo, Liquid, PictorialBar
from pyecharts import options as opts
from pyecharts.globals import ThemeType, SymbolType
from pyecharts.charts import Page, Grid
import xlrd
from pyecharts.globals import ChartType
import datetime

app = Flask(__name__, static_folder="templates")
# 获取当前的时间
time_ = datetime.datetime.now().strftime("%Y年%m月%d日-%H时%M分")
print(time_)
"""
1.数据源的准备及处理
使用excel表格中的数据
"""
try:
    data_table = xlrd.open_workbook(r"模拟网络攻击数据.xls")
except Exception as e:
    print(e)
# xlrd库的基本使用方法：-----读表格中的数据
# print(data_table.sheet_names())#获取所有的表名
# print(data_table.sheet_names()[0])#获取第一个表的表名
# print(data_table.sheet_by_index(0).ncols)#获取第一个表的列数
# print(data_table.sheet_by_index(0).nrows)#获取第一个表的行数
# print(data_table.sheet_by_index(0).cell(3,0).value)#获取第1个表的第4行第1列的数据
# print(data_table.sheet_by_index(0).col_values(0))#获取第1列的所有数据
# print(data_table.sheet_by_index(0).row_values(2))#获取第三行的所有数据
# for row in range(3,data_table.sheet_by_index(0).nrows):#获取第4行开始的第一列的数据
#     print(data_table.sheet_by_index(0).cell_value(row,0))

'''攻击线路图数据
计算：追加
展示:所有值'''
geo_data = data_table.sheet_by_index(1).cell_value(0, 0)  # 图表标题
geo_data_from_province_list = []  # 所有发起攻击的省份list
geo_data_to_province_list = []  # 所有被攻击的省份list
geo_data_to_province_num_list = []  # 省份被攻击的数量
for row in range(2, data_table.sheet_by_index(1).nrows):
    geo_data_from_province_list.append(data_table.sheet_by_index(1).cell_value(row, 0))
    geo_data_to_province_list.append(data_table.sheet_by_index(1).cell_value(row, 1))
    geo_data_to_province_num_list.append(data_table.sheet_by_index(1).cell_value(row, 2))
geo_from_to_list = list(zip(geo_data_from_province_list, geo_data_to_province_list))  # 组合成从发起到目的地的二维list
geo_to_num_list = list(zip(geo_data_from_province_list, geo_data_to_province_num_list))  # 组合成从被攻击的省份,攻击数的二维list

'''来自省份的攻击数
计算：追加
展示：所有数据'''
province_title = data_table.sheet_by_index(1).cell_value(0, 0)  # 图表标题
list_province = []  # 省份名字list
for row in range(2, data_table.sheet_by_index(1).nrows):
    list_province.append(data_table.sheet_by_index(1).cell_value(row, 0))
print(list_province)
list_province_num = []  # 攻击次数list
for row in range(2, data_table.sheet_by_index(1).nrows):
    list_province_num.append(data_table.sheet_by_index(1).cell_value(row, 1))
print(list_province_num)  # [省份,攻击次数]两个一维list变成一个二维list
from_province = list(zip(list_province, list_province_num))
print(from_province)

'''被攻击的对象
计算方式：追加
展示：所有值'''
obj_title = data_table.sheet_by_index(2).cell_value(0, 0)  # 图表标题
obj_list = []  # 被攻击的对象类型名称
for row in range(2, data_table.sheet_by_index(2).nrows):
    obj_list.append(data_table.sheet_by_index(2).cell_value(row, 0))
# print(obj_list)
obj_num_all_list = []  # 对象被攻击的总次数
for row in range(2, data_table.sheet_by_index(2).nrows):
    obj_num_all_list.append(data_table.sheet_by_index(2).cell_value(row, 1))
print(obj_num_all_list)
obj_num_sccess_list = []  # 对象被攻击的成功的次数
for row in range(2, data_table.sheet_by_index(2).nrows):
    obj_num_sccess_list.append(data_table.sheet_by_index(2).cell_value(row, 2))
print(obj_num_sccess_list)
col1_name = data_table.sheet_by_index(2).cell_value(1, 1)  # 数据项1的图例名称
col2_name = data_table.sheet_by_index(2).cell_value(1, 2)  # 数据项2的图例名称

'''被攻击的方式
计算方式：追加
展示：所有数据'''
tool_title = data_table.sheet_by_index(3).cell_value(0, 0)  # 图表标题
tool_list = []  # 攻击方式名称
for row in range(2, data_table.sheet_by_index(3).nrows):
    tool_list.append(data_table.sheet_by_index(3).cell_value(row, 0))
# print(obj_list)
tool_num_all_list = []  # 此方式攻击的总次数
for row in range(2, data_table.sheet_by_index(3).nrows):
    tool_num_all_list.append(data_table.sheet_by_index(3).cell_value(row, 1))
print(tool_num_all_list)
tool_num_sccess_list = []  # 此方式攻击的成功的次数
for row in range(2, data_table.sheet_by_index(3).nrows):
    tool_num_sccess_list.append(data_table.sheet_by_index(3).cell_value(row, 2))
print(tool_num_sccess_list)
tool_col1_name = data_table.sheet_by_index(3).cell_value(1, 1)  # 数据项1的图例名称
tool_col2_name = data_table.sheet_by_index(3).cell_value(1, 2)  # 数据项2的图例名称

'''成功与失败的占比
计算:追加
展示:最后一次数据'''
pie_title = data_table.sheet_by_index(4).cell_value(0, 0)  # 图表标题
name_pie_failed = data_table.sheet_by_index(4).cell_value(1, 1)  # 数据项1的图例名称
name_pie_sccess = data_table.sheet_by_index(4).cell_value(1, 2)  # 数据项2的图例名称
data_pie_failed = data_table.sheet_by_index(4).cell_value(-1, 1)  # 数据项1的数值
data_pie_sccess = data_table.sheet_by_index(4).cell_value(-1, 2)  # 数据项2的数值
list_failed_sccess = [[name_pie_failed, data_pie_failed], [name_pie_sccess, data_pie_sccess]]  # 组成一个二维的list
print(list_failed_sccess)

'''词云图的数据
计算：追加
展示：所有值'''
wordcloud_title = data_table.sheet_by_index(5).cell_value(0, 0)  # 词云图的图表标题
word_list = []  # word词_list
for row in range(2, data_table.sheet_by_index(5).nrows):
    word_list.append(data_table.sheet_by_index(5).cell_value(row, 0))
word_num_list = []  # word数量_list
for row in range(2, data_table.sheet_by_index(5).nrows):
    word_num_list.append(data_table.sheet_by_index(5).cell_value(row, 1))
wordcloud_list = list(zip(word_list, word_num_list))  # 将word和对应数量组成一个二维list
print(wordcloud_list)

'''折线图数据
计算：追加
展示：最后5个数据'''
line_title = data_table.sheet_by_index(6).cell_value(0, 0)  # 折线图的图表标题
line_x_list = []  # X轴,最新5个值
for row in range(data_table.sheet_by_index(6).nrows - 8, data_table.sheet_by_index(6).nrows):
    line_x_list.append(data_table.sheet_by_index(6).cell_value(row, 0))
line_y_name = data_table.sheet_by_index(6).cell_value(1, 1)  # y轴对应的数据项名称
line_y_list = []  # y轴数据_list，最新5个值
for row in range(data_table.sheet_by_index(6).nrows - 8, data_table.sheet_by_index(6).nrows):
    line_y_list.append(data_table.sheet_by_index(6).cell_value(row, 1))

'''象形柱状图数据：
计算：top6
展示：top6'''
pic_x_data = []
pic_x_value_data = []
for row in range(2, data_table.sheet_by_index(1).nrows):
    pic_x_data.append(data_table.sheet_by_index(1).cell_value(row, 0))
    pic_x_value_data.append(data_table.sheet_by_index(1).cell_value(row, 2))
pic_list = list(zip(pic_x_data, pic_x_value_data))
print(pic_list[0][0])
new_list = sorted(pic_list, key=(lambda x: x[1]), reverse=True)  # 排序
print(new_list)
top6_list = []  # 排前6的list
for i in range(0, 6):
    top6_list.append(new_list[i])
print(top6_list)
top_x = []
for j in range(0, 6):
    top_x.append(top6_list[j][0])
print(top_x)
top_v = []
for z in range(0, 6):
    top_v.append(top6_list[z][1])
print(top_v)
"""
2.插入图表
"""
'''1.攻击线路图
计算：追加
显示：所有值'''

geo_1 = (
    Geo(init_opts=opts.InitOpts(bg_color="rgba(39,6,195,1)"))
        .add_schema(

        maptype="china",
        itemstyle_opts=opts.ItemStyleOpts(),

    )
        .add(
        "",
        geo_to_num_list,  # 到达的城市
        type_=ChartType.EFFECT_SCATTER,
        color="",
    )

        .add(  # 攻击路径
        "",
        geo_from_to_list,
        type_=ChartType.LINES,
        effect_opts=opts.EffectOpts(
            symbol=SymbolType.ARROW, symbol_size=10, color="blue"
        ),
        linestyle_opts=opts.LineStyleOpts(curve=0.1, color=""),

    )
        .set_series_opts(label_opts=opts.LabelOpts(is_show=False))
        .set_global_opts(
        # legend_opts=opts.LegendOpts(is_show=True),
        title_opts=opts.TitleOpts(
            title="网络攻防监控",
            pos_left="center",
            pos_top="2%",
            title_textstyle_opts=opts.TextStyleOpts(
                font_size=80,
                color="white"

            ),
            subtitle="\n" + time_,
            subtitle_textstyle_opts=opts.TextStyleOpts(
                font_size=39,
                color="darkgreen",
                font_style="bold"

            )

        )
    )
)

'''图3-柱状图：各个对象的被攻击的次数'''
bar_1 = (  # 实例化Bar图表
    Bar(init_opts=opts.InitOpts(theme=ThemeType.INFOGRAPHIC))  # 设置图表的主题
        .add_xaxis(obj_list)  # 柱状图横坐标
        .add_yaxis(col1_name, obj_num_all_list)  # 柱状图纵坐标图例名称+数据项
        .add_yaxis(col2_name, obj_num_sccess_list)  # 柱状图纵坐标图例名称+数据项
        .set_global_opts(
        title_opts=opts.TitleOpts(title=obj_title, pos_top="3%", pos_left="7%",
                                  title_textstyle_opts=opts.TextStyleOpts(color="white")),  # 设置图表标题
        legend_opts=opts.LegendOpts(is_show=True,
                                    pos_left="6%",
                                    pos_top="33%",
                                    textstyle_opts=opts.TextStyleOpts(
                                        font_size=25,
                                        color="yellow"
                                    )
                                    ),  # 图例属性
        xaxis_opts=opts.AxisOpts(  # X轴属性(名称+字的大小)
            name="",
            name_textstyle_opts=opts.TextStyleOpts(
                font_size=40,
                color="black"
            ),
            axislabel_opts=opts.LabelOpts(
                color="white",
                font_size=12,
                font_style="bold"
            ),
            axistick_opts=opts.AxisTickOpts(is_align_with_label=True),  # 坐标轴刻度配置项
            axisline_opts=opts.AxisLineOpts(linestyle_opts=opts.LineStyleOpts(
                width=2,
            )
            ),  # 轴线配置项

        ),
        yaxis_opts=opts.AxisOpts(  # y轴属性(名称+字的大小)
            name="",
            name_textstyle_opts=opts.TextStyleOpts(
                font_size=30,
                color="black"
            ),
            axislabel_opts=opts.LabelOpts(
                font_size=17,
            ),
            axistick_opts=opts.AxisTickOpts(is_align_with_label=True),  # 坐标轴刻度配置项
            axisline_opts=opts.AxisLineOpts(linestyle_opts=opts.LineStyleOpts(
                width=2,
                # color="black"
            )
            ),  # 轴线配置项

        ),

    )
        .set_series_opts(label_opts=opts.LabelOpts(is_show=True,
                                                   font_size=20,
                                                   color="white"
                                                   )
                         )  # 是否显示值标签

)
'''图4-柱状图：各种方式攻击的次数'''
bar_2 = (  # 实例化Bar图表
    Bar()  # 设置图表的主题
        .add_xaxis(tool_list)  # 柱状图横坐标
        .add_yaxis(tool_col1_name, tool_num_all_list)  # 柱状图纵坐标图例名称+数据项
        .add_yaxis(tool_col2_name, tool_num_sccess_list)  # 柱状图纵坐标图例名称+数据项
        .set_global_opts(
        title_opts=opts.TitleOpts(title=tool_title,
                                  pos_left="80%",
                                  pos_top="3%",
                                  title_textstyle_opts=opts.TextStyleOpts(
                                      color="white",
                                      font_size=40,
                                      font_style="bold"
                                  )
                                  ),
        legend_opts=opts.LegendOpts(is_show=True,
                                    pos_left="82%",
                                    pos_top="33%",
                                    textstyle_opts=opts.TextStyleOpts(
                                        font_size=25,
                                        color="yellow"
                                    )
                                    ),  # 图例属性
        xaxis_opts=opts.AxisOpts(  # X轴属性(名称+字的大小)
            name="",
            name_textstyle_opts=opts.TextStyleOpts(
                font_size=40,
                color="black"
            ),
            axislabel_opts=opts.LabelOpts(
                color="white",
                font_size=9,
                font_style="bold"
            ),
            axistick_opts=opts.AxisTickOpts(is_align_with_label=True),  # 坐标轴刻度配置项
            axisline_opts=opts.AxisLineOpts(linestyle_opts=opts.LineStyleOpts(
                width=2,
            )
            ),  # 轴线配置项

        ),
        yaxis_opts=opts.AxisOpts(  # y轴属性(名称+字的大小)
            name="",
            name_textstyle_opts=opts.TextStyleOpts(
                font_size=30,
                color="black"
            ),
            axislabel_opts=opts.LabelOpts(
                font_size=17,
            ),
            axistick_opts=opts.AxisTickOpts(is_align_with_label=True),  # 坐标轴刻度配置项
            axisline_opts=opts.AxisLineOpts(linestyle_opts=opts.LineStyleOpts(
                width=2,
                # color="black"
            )
            ),  # 轴线配置项

        ),

    )
        .set_series_opts(label_opts=opts.LabelOpts(is_show=True,
                                                   font_size=20,
                                                   color="white"
                                                   )
                         )  # 是否显示值标签

)
'''图5-饼图：被攻击成功与失败的占比'''  # 此处需要使用隐形的Bar图表，适应Grid布局，系列配置项以bar的配置为主


def pie_zuhe():
    pie_1 = (  # 图表实例化
        Pie(init_opts=opts.InitOpts(theme=ThemeType.WONDERLAND))  # 设置图表标题
            .add('', list_failed_sccess,
                 radius=["5%", "10%"],  # 设置图表的内网半径
                 center=["28%", "17%"])  # 设置图表的布局位置
            .set_colors(["white", "yellow"])  # 颜色
            .set_global_opts(  # 全局配置：设置标题
            title_opts=opts.TitleOpts(title=pie_title, pos_left="center",
                                      title_textstyle_opts=opts.TextStyleOpts(color="white")),
            legend_opts=opts.LegendOpts(is_show=True)

        )
            .set_series_opts(  # 图表内部配置：
            label_opts=opts.LabelOpts(
                formatter="{b}\n{c}",  # 标签格式
                is_show=True,  # 是否显示标签
                font_size=20,
                color="yellow"

            )
        )
    )
    bar = (  # 隐形的Bar图表
        Bar(init_opts=opts.InitOpts(width="300px", height="200px"))  # 设置图表的主题
            .add_xaxis([""])
            .add_yaxis("", [""])
            .set_global_opts(
            xaxis_opts=opts.AxisOpts(
                axisline_opts=opts.AxisLineOpts(
                    linestyle_opts=opts.LineStyleOpts(
                        width=0,
                    )
                )),  # 轴线配置项
            yaxis_opts=opts.AxisOpts(  # y轴属性(名称+字的大小)
                axisline_opts=opts.AxisLineOpts(
                    linestyle_opts=opts.LineStyleOpts(
                        width=0,
                    )
                ),  # 轴线配置项

            ),
            title_opts=opts.TitleOpts(title=pie_title, pos_top="8%", pos_left="22%",
                                      title_textstyle_opts=opts.TextStyleOpts(color="white")),
            legend_opts=opts.LegendOpts(is_show=False, pos_top="40%", pos_left="20%", pos_right="70%",
                                        pos_bottom="60%"),

        )
            .set_series_opts(label_opts=opts.LabelOpts(is_show=True, font_size=20))  # 是否显示值标签

    )
    c1 = bar.overlap(pie_1)

    return c1


'''图6.词云图：解决办法汇总'''
wordcloud_1 = (  # 实例化对象
    WordCloud()  # 设置图表主题
        .add(
        series_name="",  # 设置图例名称,并没什么用
        data_pair=wordcloud_list,  # 设置数据
        word_size_range=[20, 70],  # 字体大小范围
        shape=SymbolType.ARROW,  # 样式
        pos_left="51%", pos_top="18%"
    )
        .set_global_opts(
        title_opts=opts.TitleOpts(title=wordcloud_title, pos_left="82%", pos_top="43%",
                                  title_textstyle_opts=opts.TextStyleOpts(color="white"),
                                  ),
        tooltip_opts=opts.TooltipOpts(is_show=True),
    )
)
'''图7.面积图:最新5个时间段内的被攻击数'''
line_1 = (  # 实例化图表
    Line()  # 设置图表主题
        .add_xaxis(line_x_list)  # 设置横坐标
        .add_yaxis(line_y_name,
                   line_y_list,
                   is_smooth=True
                   )  # 设置纵坐标1、突出显示最大值或者最小值

        .set_global_opts(  # 全局配置(标题、横坐标、纵坐标、图例、)
        legend_opts=opts.LegendOpts(is_show=False),
        title_opts=opts.TitleOpts(title="",
                                  pos_left="1%",
                                  pos_top="75%",
                                  title_textstyle_opts=opts.TextStyleOpts(
                                      color="white")
                                  ),  # 标题属性
        xaxis_opts=opts.AxisOpts(  # X轴属性(名称+字的大小)
            name="",
            name_textstyle_opts=opts.TextStyleOpts(
                font_size=30,
                color="white"
            ),
            axislabel_opts=opts.LabelOpts(
                font_size=15,
                color="white"
            ),
            axistick_opts=opts.AxisTickOpts(is_align_with_label=True),  # 坐标轴刻度配置项
            axisline_opts=opts.AxisLineOpts(linestyle_opts=opts.LineStyleOpts(
                width=3,
            )
            ),  # 轴线配置项
            is_scale=False,  # 是否按比例收缩
            boundary_gap=False,  # 是否两侧空白
        ),
        yaxis_opts=opts.AxisOpts(  # y轴属性(名称+字的大小)
            name="",
            name_textstyle_opts=opts.TextStyleOpts(
                font_size=30,
                color="black"
            ),
            axislabel_opts=opts.LabelOpts(
                font_size=15,
            ),
            axistick_opts=opts.AxisTickOpts(is_align_with_label=True),  # 坐标轴刻度配置项
            axisline_opts=opts.AxisLineOpts(linestyle_opts=opts.LineStyleOpts(
                width=3,
            )
            ),  # 轴线配置项
            is_scale=False,  # 是否按比例收缩
            boundary_gap=False,  # 是否两侧空白
        ),

    )
        .set_series_opts(
        areastyle_opts=opts.AreaStyleOpts(opacity=0.5),  # 颜色区别度
        label_opts=opts.LabelOpts(is_show=True, font_size=20),  # 是否显示标签值
        linestyle_opts=opts.LineStyleOpts(color="red", width=2)  # 设置线条属性

    )

)

'''图10-象形柱状图'''
pic_1 = (
    PictorialBar()
        # .add_xaxis(geo_data_from_province_list)
        .add_xaxis(top_x)
        .add_yaxis(
        "",
        top_v,

        label_opts=opts.LabelOpts(is_show=True),
        symbol_size=18,
        symbol_repeat="fixed",
        symbol_offset=[0, 0],
        is_symbol_clip=True,
        symbol=SymbolType.RECT,
        color="red"
    )
        .reversal_axis()
        .set_global_opts(
        title_opts=opts.TitleOpts(title="攻击来源TOP6",
                                  pos_left="3%",
                                  pos_top="45%",
                                  title_textstyle_opts=opts.TextStyleOpts(
                                      color="white")
                                  ),  # 标题属性
        xaxis_opts=opts.AxisOpts(is_show=False),
        yaxis_opts=opts.AxisOpts(
            axistick_opts=opts.AxisTickOpts(is_show=False),
            axislabel_opts=opts.LabelOpts(
                font_size=30,
                color="white"
            ),
            axisline_opts=opts.AxisLineOpts(
                linestyle_opts=opts.LineStyleOpts(opacity=0)
            ),
        ),
    )
        .set_series_opts(label_opts=opts.LabelOpts(is_show=False,
                                                   font_size=20,
                                                   color="white"
                                                   )
                         )  # 是否显示值标签

)
"""
3.生成html文件，并按固定化格式装换成html文件
"""


def show_web():
    web_grid1 = (
        Grid(init_opts=opts.InitOpts(
            # width="2850px",
            # height="1600px",
            width="1920px",
            height="1080px",
            bg_color="rgba(39,6,195,1)"
            # bg_color="black"
        )
        )

            .add(bar_1, grid_opts=opts.GridOpts(pos_left="2%", pos_right="80%", pos_bottom="70%", pos_top="5%"))
            .add(bar_2, grid_opts=opts.GridOpts(pos_left="74%", pos_right="0%", pos_bottom="70%", pos_top="5%"))
            .add(pic_1, grid_opts=opts.GridOpts(pos_left="5%", pos_right="80%", pos_bottom="20%", pos_top="50%"))
            .add(pie_zuhe(), grid_opts=opts.GridOpts())
            .add(line_1, grid_opts=opts.GridOpts(pos_left="3%", pos_right="3%", pos_top="85%"))
            .add(geo_1, grid_opts=opts.GridOpts())
            .add(wordcloud_1, grid_opts=opts.GridOpts())

    )
    return web_grid1


@app.route("/")
def index():
    c = show_web()
    return Markup(c.render_embed())


if __name__ == '__main__':
    app.run(host="0.0.0.0")
