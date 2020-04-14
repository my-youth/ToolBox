ToolBox是业余学习VBA编写的Excel插件，整合了部分网友的代码，也新写了很多功能。
<!-- more -->

从2018年初开始学习VBA，学习过程中，不断遇到实际需要解决的问题，于是萌生了编写一个工具的想法，ToolBox就此诞生了。此过程中参考借鉴了很多网友的代码，也整合了进来，现在全部释放出来，仅供参考使用。
![](file://D:/文档/Gridea/post-images/1576918344716.png)

>下载地址：[Github](https://github.com/my-youth/ToolBox)   [百度网盘](https://pan.baidu.com/s/14coXQLpKg1moJCYvOfzMGg)

<hr style="height:10px;border:none;border-top:10px groove skyblue;" />

## 安装方法：
1. 将ToolBox.xlam文件复制到Office加载项文件夹：%userprofile%\AppData\Roaming\Microsoft\AddIns；
   ![](file://D:/文档/Gridea/post-images/1576918597308.png)
2. 回到Excel中，启用开发工具选项卡，在Excel加载项中勾选ToolBox即可使用。
   ![](file://D:/文档/Gridea/post-images/1576918658413.png)

## 更新日志：
1. 2020/4/14     新增：增加了工作表和工作簿解密功能，工作簿解密功能会丢失VBA工程数据。
2. 2020/2/2       新增：更新并完善了显示设置，可以方便地批量设置显示外观等。

## 主要功能
### 1. 选项卡功能
   1. 常用功能。整合Excel各个选项卡中常用的命令，如清除、冻结窗格、排序和筛选、数据透视向导等；
   ![常用功能](file://D:/文档/Gridea/post-images/1576917720096.png)
   1. BOE工具。是针对BOE开发的方便工作的工具。除了跳转命令外， 主要有加权Cycle Time的计算、字符合并的Plus版本和基板切割计算；
   ![BOE工具](file://D:/文档/Gridea/post-images/1576917748563.png)
   1. 数据处理。主要有区域转图、字符的合并和提取、表单处理、单元格处理、VBA功能的加解密；
   ![数据处理](file://D:/文档/Gridea/post-images/1576917783476.png)
   1. 其他小工具。有从图表标签转移过来的取色器、二维码功能、翻译插件、个人所得税计算、进制转换、词频分析、批量插入图片、文件树等。
   ![工具集](file://D:/文档/Gridea/post-images/1576917899481.png)
   1. 万能输入表。整合了[拉小登](http://www.ladeng6666.com/blog/)的万能输入表，可以将二维表快速转换为一维表。
   ![万能输入表](file://D:/文档/Gridea/post-images/1576918123014.png)
### 2. 自定义函数
   根据需求，编写了很多Function，包含“Y_”开头的函数和“P_”开头的函数，其中“P_”来自[拉小登](http://www.ladeng6666.com/blog/)。
   ![Y_函数](file://D:/文档/Gridea/post-images/1576918279721.png)
