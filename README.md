# Transform-Plug-in

/*
2.0Update:     增加<是否标红>功能，标红输入数字1即可。
*/

/*
1.1Update:     换牌、周保、月保的数据显示已经去除。
*/

这是转换插件的使用说明：

首先说明，因为缺少全面的测试，这个插件很可能存在超多bug <笑哭>，敬请见谅。
（然后计划表上不同订单的颜色区分功能我没做出来 <水平不行> ）

把压缩包解压之后文件夹含有Excel文件data.xlsx和文件夹Transform Plug-in。
data.xlsx是这个插件的输入文件，同时也是开发时的测试数据（乱写的）。

暂时不支持win7使用



使用方法：

将data.xlsx拷贝到Transform Plug-in文件夹中的dist文件夹中，
dist文件夹内有Transform.exe，双击它，
可以看到dist文件夹中生产了所需的表格。
注意生成新表格之前需要将原先生成的表格关闭或者删除。



注意事项：

这个插件的输入文件必须是名为data.xlsx的Excel文件（Microsoft Office EXCEL 2007+）。
直接修改data.xlsx，之后双击Transform.exe，就可以生产新的表格（可能生成的文件名不变但实际的内容已经改动）。

生成新表格之前需要将原先生成的表格关闭或者删除。

关于data.xlsx的说明：
表头的栏目分别是：
机器      产品名称      工单号      总量（万个）      效率（万个/班）      开始日期
表头不能被修改，否则程序无法正常运行。

表格非内容区域不能有其他数据或者字符。

数据表格的字体颜色以及单元格背景颜色可以任意修改，不会影响程序运行。

订单的排序严格按照从上到下的顺序：即对于一个机器的所有订单，会严格按照输入数据从上到下的优先级进行排班。

输入表格中不能有空行。

输入订单的时候需要保证机器是连续的，即先输入某一台机器的所有订单，再输入另外一台机器的订单。

一天按三班计算，周日全部设成休息。

可以指定订单的开工日期（即开始日期），且每一台机器的第一个订单必须指定开工日期。
后续订单开工日期可以为空，表示紧接上一订单开工。
指定中间订单的开工时间表明本强制订单到某一日期才开始，如果前面累积的订单所需要的完成时间已经超过这个日期，则当前订单
可能会自动后延或者被吃掉（找不到这个订单）

由于最终结果保留两位小数，所以可能存在小数点后两位的精度误差。

本插件在写程序的时候其实并没有考虑<换牌>以及<周/月保>。
但是后期发现也许可以把<周/月保>和<换牌>当成一个订单来处理（大概可以？不知道有没有问题）
比如对某台机器<换牌>可以当作该机器的一个总量为0.5，效率为1的订单，工单号和开始日期留空，夹在两个产品中间。
<周/月保>也许也可以根据时长来设计。
