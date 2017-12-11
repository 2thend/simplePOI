#### 一、simplePOI 项目介绍
##### simplePOI项目是基于apache POI项目对Excel导出／导入进行对封装，目的是为了简化开发者对使用，开发者仅需数行代码即可实现Excel导出功能。
##### 项目核心类只有四个：
###### ExportExcelUtils：Excel导出工具类
###### ISheetBuilder：抽象sheet操作借口
###### BaseSheetBuilder：基础操作sheet实现类
###### ExportInfo：导出Excel数据实体

##### 自定义扩展实现参考类
###### com.thend.master.exporter.CustomSheetBuilder
###### 使用装饰者模式，继承ISheetBuilder，内部持有BaseSheetBuilder引用

##### 项目启动 MyApplication.main
##### 项目访问 http://localhost:8080

#### 二、simplePOI 设计思想
##### 1、针对变化进行定制
###### 查看POI操作Excel流程不难发现，(大的模块)整体流程针对一个workbook对象和一个或多个Sheet对象进行操作，且对sheet操作为主体操作。
###### 我将对sheet操作可能出现的变化抽象出6个方法(buildTitleStyle、buildDataStyle、buildFootStyle、buildTitle、buildData、buildFoot)

##### 2、仿照Jquery的链式调用
###### 在ISheetBuilder接口中，对titile、data、footData数据操作的方法返回ISheetBuilder接口本身，实现链式调用

##### 3、单sheet定制使用装饰者模式，避免类的继承

##### 4、线程安全
###### 全局的workbook在ExportExcelUtils内部使用ThreadLocal维护
###### Sheet定制的相关信息通过ISheetBuilder实现类的构造函数，以实例化的方式将信息以ExportInfo对象的形式存放在ISheetBuilder实现类中
###### ExportExcelUtils(workbook)与ISheetBuilder实现类(Sheet)的数据通过ISheetBuilder接口中getExportInfo方法进行数据通信

#### 三、simplePOI 项目特色
##### 1、简单易用
##### 2、面向ISheetBuilder接口编程
##### 3、ISheetBuilder实现类有类似Jquery的链式调用
##### 4、单sheet定制使用装饰者模式，避免类的继承
##### 5、线程安全

>##### 项目地址：
###### github: https://github.com/2thend/simplePOI
###### 码云：https://gitee.com/ktshws/simplePOI.git
