基于python的pyecharts模块实现将Excel中的数据可视化案例——网络攻防监控
当数据分析遇上数据可视化时
![image](https://user-images.githubusercontent.com/46338963/153325223-11bdb973-a704-4336-8444-501544946dfc.png)
支持：多类型图表、多数据类型、多web框架
图片只要是数据，就有办法可视化

1.概要
1.1 题目：基于python的pyecharts模块实现将excel中的数据可视化案例—网络攻防监控
1.2 web框架：Falsk
数据刷新方式：Excel中的数据发生改变后需要重新运行此app.py文件
1.3 用到的包：pyecharts、xlrd、flask、datatime
1.4 步骤：
------001：导入相关的包
------002：数据源的准备及处理
------003：作图,图表的相关属性设置
------004：图表组合，Grid布局
------005：flask（app.route）
------006：run app.py

2.文件目录结构
![image](https://user-images.githubusercontent.com/46338963/153325112-def40456-f945-45ad-b658-9c6f47f18dc5.png)
pyecharts离线包地址(templates):
https://github.com/pyecharts/pyecharts/archive/refs/heads/master.zip

3.如何运行
表格中的数据发生改变都需要重新执行这三个步骤
3.1更改Excel表格中的数据
3.2python3 app.py
3.3浏览器访问：ip:5000

4.运行效果
![image](https://user-images.githubusercontent.com/46338963/153323422-304478c2-d2e8-44b4-bf93-4a435befe82f.png)
