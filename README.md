# AnalysisTable-inDocxPdf-byPython    
[Python解析docx/pdf文件中表格内容](https://www.yuque.com/yolanda-7ksta/fplovr/rclbtivc91ls5zfg/edit)

###  项目背景  
每个团体可能都有招聘工作，招聘过程中会接收非常多简历。部分公司如学校/事业单位类会采取报名机制，需要下载表格填写相关个人信息，然后学校内部根据基础信息表进行筛选。

在后期，单位想汇总某个时期所有应聘者相关信息进行数据分析。因此想将个人基础信息表中所关注的信息汇总到Excel表格中。若手动复制粘贴将是非常大的工作量，且每年都会有重复类型工作，因此靠人力非常低效。  

（若能考虑搭建线上报名系统，问题都能迎刃而解。此项目建立在没有线上报名系统的前提上。）  


### 应用场景
> 将大量含同格式不规则表格的文档中部分内容汇总到一张Excel表格中。


### 解决方案
本项目拟**借助Python中docx和pdfplumber模块来解析docx/pdf文件中表格内容并汇总**。

注意事项：

● 文档可分散在不同文件夹，可实现遍历。

● Python中解析doc文件较麻烦，因此文档尽可能采用docx/pdf格式或提前将doc转化为docx/pdf格式。

### 项目代码
代码实现：table_file_tool.py
测试文件夹：test_file
