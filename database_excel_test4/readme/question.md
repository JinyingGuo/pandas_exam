# 一、不修改程序，怎么能读不同的数据文件
    通过os模块获取符合条件的数据文件。并通过控制台输出，可根据序号选择需要处理哪些表格。
   ![question1](../readme/pictures/question1.png)
    
# 二、测试数据应包含正确数据，也要包含不一致数据的组合，每个同学要设计不同数据的组合，测试你的程序
## 1.数据库原理及应用课程目标达成情况-updata1.xls表格
    （1）**修改**：增加了课程目标5；删除了达成度（教师评价法）
    （2）**测试结果**：
        程序正常运行并计算出结果。
        数据正确性测试运行通过。数据有不一致情况，主要因为精度问题导致不一致。
   ![updata1测试1](../readme/pictures/update1-1.png)
   ![updata1测试2](../readme/pictures/update1-2.png)
   ![updata1测试3](../readme/pictures/update1-3.png)
   ![updata1测试4](../readme/pictures/update1-4.png)

## 2.数据库原理及应用课程目标达成情况-updata2.xls表格
    （1）**修改**：增加了课程目标5；增加了编程考核方式；增加了4班；删除了达成度（教师评价法）
    （2）**测试结果**：
        程序正常运行并计算出结果。
        数据正确性测试程序运行通过。数据有不一致情况，主要因为精度问题导致不一致。
   ![updata2测试1](../readme/pictures/update2-1.png)
   ![updata2测试2](../readme/pictures/update2-2.png)
   ![updata2测试3](../readme/pictures/update2-3.png)
   ![updata2测试4](../readme/pictures/update2-4.png)
## 3.数据库原理及应用课程目标达成情况-updata3.xls表格
    （1）**修改**：修改了考试页帧得分和达成度数据
    （2）**测试结果**：
   ![updata3测试1](../readme/pictures/update3-1.png)
   ![updata3测试2](../readme/pictures/update3-2.png)
   ![updata3测试3](../readme/pictures/update3-3.png)
   ![updata3测试4](../readme/pictures/update3-4.png)
   ![updata3测试4](../readme/pictures/update3-5.png)
   ![updata3测试4](../readme/pictures/update3-6.png)
   ![updata3测试4](../readme/pictures/update3-7.png)
# 三、测试不同数据情况，学习使用log日志的用法
    (1)使用loguru记录日志。
    (2)创建excel文件同名log日志
       为每一个处理的excel表格创建同名log日志，可具体查看某一个文件的处理结果。
    (3)使用debug、info、error、exception记录程序运行情况
   ![loguru](../readme/pictures/log1.png)
   ![loguru](../readme/pictures/log2.png)
   ![loguru](../readme/pictures/log3.png)
   ![loguru](../readme/pictures/log4.png)
   ![loguru](../readme/pictures/log5.png)
   

# 四、每位同学测试其他10位同学的测试数据，记录程序出现的问题，并进行修改。
##1.李善勤同学表格
###（1）
   ![问题1](../readme/pictures/李善勤表格问题1.png)
   ![错误](../readme/pictures/李善勤问题1错误.png)
   
   权重值获取失败：修改程序后可正确读取
### (2)
   ![问题2](../readme/pictures/李善勤表格问题2.png)
   
   考核方式个数不一致。导致读取页帧个数出现问题
   修改了主观评价页帧个数获取方式，已解决
## 2.索佳峰同学表格
   ![问题](../readme/pictures/索佳峰表格问题.png)
   ![错误](../readme/pictures/索佳峰表格出现错误.png)
   
   计算达成度时，被除数出现0值导致程序无法运行
## 3.张振伟同学表格
   ![错误](../readme/pictures/张振伟表格问题.png)
   
   删除了考试页帧，导致程序遍历了所有页帧也无法找到班级列而无法运行