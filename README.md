# tableExtractor
Extract specific table from doc or docx files to CSV file

主要用了POI包

对于DOC文件，逐段落进行扫描，直到扫描到自己所需要的段落特征，然后通过检查该段落在表格中的状况，拼凑出字符串然后输出到CSV格式

对于DOCX文件，直接获得表格迭代器，迭代直到获得自己想要的表格，然后拼凑字符串导出根据自己表格的不同特征需要对代码中进行修改
