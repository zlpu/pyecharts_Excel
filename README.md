<h1 align = "center">当数据分析遇上数据可视化时</h1>

![在这里插入图片描述](https://img-blog.csdnimg.cn/838d477a3d6c4d80bbd471daca44ad12.png)
支持：多类型图表、多数据类型、多web框架


只要是数据，就有办法可视化

<h3 align = "center">1.概要</h3>
01.题目：基于python的pyecharts模块实现将excel中的数据可视化案例—网络攻防监控

02.web框架：Falsk
数据刷新方式：Excel中的数据发生改变后需要重新运行此app.py文件

03.用到的包：pyecharts、xlrd、flask、datatime

04.excel中的数据的更新方式：
修改数据或者新增行数据

05.步骤：

------001：导入相关的包

------002：数据源的准备及处理

------003：作图,图表的相关属性设置

------004：图表组合，Grid布局

------005：flask（app.route）

------006：run app.py

<h3 align = "center">2.文件目录结构</h3>

![在这里插入图片描述](https://img-blog.csdnimg.cn/52cc72f47155442e8373adf6f3687749.png?x-oss-process=image/watermark,type_d3F5LXplbmhlaQ,shadow_50,text_Q1NETiBASVTlsI_lk6Xlk6U=,size_13,color_FFFFFF,t_70,g_se,x_16)

pyecharts离线包地址(templates):
[https://github.com/pyecharts/pyecharts/archive/refs/heads/master.zip](https://github.com/pyecharts/pyecharts/archive/refs/heads/master.zip)

<h3 align = "center">3.如何运行</h3>

表格中的数据发生改变都需要重新执行这三个步骤

01.更改Excel表格中的数据

02.python3 app.py

03.浏览器访问：ip:5000

<h3 align = "center">4.运行效果</h3>
![在这里插入图片描述](https://img-blog.csdnimg.cn/12242146367e4bb79b4b9639a83dc1dd.png?x-oss-process=image/watermark,type_d3F5LXplbmhlaQ,shadow_50,text_Q1NETiBASVTlsI_lk6Xlk6U=,size_20,color_FFFFFF,t_70,g_se,x_16)

