# Bond-Trade-System
A project of the class Bond Trade System
Developed by VBA in excel 2010

Dcription in Chinese:
本次porject得成果集中于文件bond.trading.system.project(魏翔宇).xls中。

文件中存在四个sheet
sheet1名为Requirement，是原有的sheet，只在第一个单元格里给出了题目要求。
sheet2名为Yield.Curve。按钮Load Yield Curve可用于选择并加载文件名为yc.2013.MM.DD.csv
的文件，并将数据填入A-L列中，同时根据所有数据绘制出yield曲线图。
sheet3名为Bond.Position。按钮Load Position可用于选择并加载文件名为bond.position.2013.
MM.DD.csv格式的文件，并将数据填入A-E列中。按钮Calculate & Plot!可用于根据所填数据，计算
出每一行债券对应的Clean Price, Accrued Interest, Dirty Price, Modified Duration, Po-
-sition BPV等数据，并绘制出四种债券的Position BPV走势图，并粘贴到sheet4中。
sheet4名为Position.BPV.Curve，sheet3中计算出的Position BPV走势图在这里显示。
代码大部分都是直接编写，绘制曲线图部分主要通过录制宏并稍加修改的方式写出。

bond.trading.system.project(魏翔宇).xls文件中已经加载了所有数据并进行了计算和绘图，为了
测试可以将所有数据和图表删除后，重新导入数据，并点击按钮查看结果。


If in need, Please see Sheet1.cls and Sheet2.cls to view the Visual Basic Code.
