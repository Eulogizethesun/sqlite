# sqlite

编写一个命令行程序，它具有三个命令行参数：
1.	数据库名
2.	XLSX文件名，带.XLSX后缀
3.	Excel中中的页名，当这个参数省略时，取第一个页
4.	数据库表名，当这个参数省略时，用页名或当第三个参数也省略时用XLSX文件名前缀作为表名  
当然第3个参数省略时不能单独带有第四个参数。  
程序要打开xlsx文件，找到指定的页面，在本地的SQLite数据库中建立指定名称的表.xlsx里的第一行作为表中的字段名，字段类型由xlsx数据推断得到。推断的规则为：  
整数的为INT类型;  
浮点数为真正的类型;  
其他为字符类型，大小为所有行中最大的长度。  
然后程序将所有的行数据导入该表，并为该表建立一个PK，以表达行号。程序输出表的结构（SQL）和行的数量。  
