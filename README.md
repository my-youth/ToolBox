ToolBox是业余学习VBA编写的Excel插件，整合了部分网友的代码，也新写了很多功能。
<!-- more -->

从2018年初开始学习VBA，学习过程中，不断遇到实际需要解决的问题，于是萌生了编写一个工具的想法，ToolBox就此诞生了。此过程中参考借鉴了很多网友的代码，也整合了进来，现在全部释放出来，仅供参考使用。
![](https://s1.ax1x.com/2020/04/24/Jr09MD.png)

>下载地址：[Github](https://github.com/my-youth/ToolBox/releases)   [百度网盘](https://pan.baidu.com/s/14coXQLpKg1moJCYvOfzMGg)   [天翼云盘](https://cloud.189.cn/t/32aaAbJNbAry)

<hr style="height:10px;border:none;border-top:10px groove skyblue;" />

## 安装方法：
**重要提示：x64版本仅用于64位Office（[版本查看方式](https://support.microsoft.com/zh-cn/office/%E5%85%B3%E4%BA%8E-office%EF%BC%9A%E6%88%91%E4%BD%BF%E7%94%A8%E7%9A%84%E6%98%AF%E5%93%AA%E4%B8%AA%E7%89%88%E6%9C%AC%E7%9A%84-office%EF%BC%9F-932788b8-a3ce-44bf-bb09-e334518b8b19)），缺少方骥的汉字转拼音功能。**
1. 将ToolBox.xlam文件复制到Office加载项文件夹：%userprofile%\AppData\Roaming\Microsoft\AddIns；
   ![](https://s1.ax1x.com/2020/04/24/Jr0MLQ.png)
2. 回到Excel中，启用开发工具选项卡，在Excel加载项中勾选ToolBox即可使用。
   ![](https://s1.ax1x.com/2020/04/24/Jr0Giq.png)

## 更新日志：
1. 2020/12/6   新增：对选定或所有单元格的批注统一设置
2. 2020/8/30   新增：整合Ron de Bruin的表驱动生成QAT菜单工具
3. 2020/8/30   新增：QAT快速访问工具栏添加常用命令按钮
4. 2020/6/27   新增：整合[Excel大全_方骥](https://weibo.com/2190827182/BAeh1xbd9)的汉字转拼音功能
5. 2020/6/21    新增：增加Y_SQL2ExcelDate函数，用于将数据库中类似于01.02.2020, 00:06:13.152的日期调整为Excel可识别的格式

## 主要功能
### 1. 选项卡功能
   1. 常用功能。整合Excel各个选项卡中常用的命令，如清除、冻结窗格、排序和筛选、数据透视向导等；
   ![常用功能](https://s1.ax1x.com/2020/04/24/Jr0oFI.png)
   1. BOE工具。是针对BOE开发的方便工作的工具。除了跳转命令外， 主要有加权Cycle Time的计算、字符合并的Plus版本和基板切割计算；
   ![BOE工具](https://s1.ax1x.com/2020/04/24/Jr07fP.png)
   1. 数据处理。主要有区域转图、字符的合并和提取、表单处理、单元格处理、VBA功能的加解密；
   ![数据处理](https://s1.ax1x.com/2020/04/24/Jr0L6S.png)
   1. 其他小工具。有从图表标签转移过来的取色器、二维码功能、翻译插件、个人所得税计算、进制转换、词频分析、批量插入图片、文件树等。
   ![工具集](https://s1.ax1x.com/2020/04/24/JrBSkn.png)
   1. 万能输入表。整合了[拉小登](http://www.ladeng6666.com/blog/)的万能输入表，可以将二维表快速转换为一维表。
   ![万能输入表](https://s1.ax1x.com/2020/04/24/JrBF6U.png)
### 2. 自定义函数
   根据需求，编写了很多Function，包含“Y_”开头的函数和“P_”开头的函数，其中“P_”来自[拉小登](http://www.ladeng6666.com/blog/)。
   ![Y_函数](https://s1.ax1x.com/2020/04/24/JrBufx.png)
