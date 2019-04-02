## 《实验1：C#控制台编程》实验报告    
学院：软件学院  班级：4 班    学号：3017218164    姓名：石琦  
日期：    2019年 3 月14  日
# 一、功能概述  
1. 在控制台生成由黑白字符块组成的二维码图形；
![黑白字符块组成的二维码](https://github.com/TJUamoeba/LAB1/blob/master/QRcreate/Pictures/%E5%AD%97%E7%AC%A6%E5%9D%97%E6%8E%A7%E5%88%B6%E5%8F%B0.png)
2. 通过读取text文档，将文档内容逐行输出为二维码图片格式并存储；
![text文档转二维码控制台输出结果](https://github.com/TJUamoeba/LAB1/blob/master/QRcreate/Pictures/TXT%E6%8E%A7%E5%88%B6%E5%8F%B0.png)
3. 读取Excel表格文件，将表格每行第一列单元格内容输出为二维码图片格式并存储；
![Excel文件转二维码控制台输出结果](https://github.com/TJUamoeba/LAB1/blob/master/QRcreate/Pictures/EXCEL%E6%8E%A7%E5%88%B6%E5%8F%B0.png)
![文件夹内图片生成情况](https://github.com/TJUamoeba/LAB1/blob/master/QRcreate/Pictures/%E6%96%87%E4%BB%B6%E5%A4%B9%E8%BE%93%E5%87%BA%E7%BB%93%E6%9E%9C.png)  
# 二、项目特色  
1. 对每个字符串的输出结果进行反馈，反馈结果有：
    i. 正常输出：二维码存储地址；
    ii.异常：生成失败的字符串位置、错误类型（如：字符串过长）
2. 对命名重复的图片文件添加了防止存储覆盖的方法；
![命名重复的情况](https://github.com/TJUamoeba/LAB1/blob/master/QRcreate/Pictures/%E9%87%8D%E5%A4%8D%E5%91%BD%E5%90%8D.PNG)
3. 字符串限定规则单独作为一个类，便于修改规则，且使用了正则表达式；	   
# 三、代码总量    
336行     
# 四、工作时间  
1周
# 五、知识点总结图（Concept MAP）
![知识点图](https://github.com/TJUamoeba/LAB1/blob/master/QRcreate/Pictures/%E7%9F%A5%E8%AF%86%E7%82%B9.PNG)
# 六、结论     
1. 学习了二维码的相关知识，了解了二维码纠错等级等概念的定义；
3. 学习了C#关于TEXT文件的读写操作；
2. 学习了C#关于EXCEL文件的读写操作；     
 *其间发现了程序运行后测试EXCEL文件只能以只读方式打开的问题，发现是由EXCEL文件读取后不会自动关闭导致的，后期添加了EXCEL文件强制关闭的方法*