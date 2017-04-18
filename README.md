# ReadExcel相关说明
## 实现功能
可以实现导入Excel表中的数据到数据库中（目前仅支持xls，xlsx文件）

## 开发环境 
jdk1.8
## 所需架包
mysql-connector-java-5.0.5-bin.jar ------ mysql数据库包

poi-3.15.jar ------ poi操作所需架包（操作xls文件）

poi-ooxml-3.15.jar ------ poi操作所需架包（操作xlsx文件）
## 使用说明
1、在数据库中创建一张表。

2、新建一张Excel表，文件名和数据库中表明一样（不区分大小写），excel表中的列要和数据库中列名一样（区分大小写）

3、调用ReadExcel.newInstance(Connection connection)获取ReadExce的实例对象，所需参数为数据库连接对象

4、调用readExcel对象的readExcelToMysql(String path)方法，所需参数为excel表的路径
## 返回状态
ReadExcel.INVALIDFILEPATH --- 无效的文件路径

ReadExcel.FILETYPEERROR --- 文件类型错误(目前仅支持xls，xlsx文件)

ReadExcel.HAVENOTABLE --- 数据库中没有相对应的表

ReadExcel.COLUMNINCONSISTENCY --- 数据库中列名和excel表中列名不一致

ReadExcel.SUCCESS --- 导入成功

## 异常
excel表中的数据如果和数据库中的数据不匹配或者冲突会导致异常，可通过调用readExcel对象的getLogs（）查看所有异常信息或者调用getNewLog（）方法查看最新信息
## 其他方法
getLogs（） --- 获取所有异常日志信息

getNewLog（） --- 获取最新的一条异常日志信息
