# ExcelUtils
### _A convenient Excel reading and writing util_
#### 一、简介：
> ##### ExcelUtils是一个参考了 [xxl-excel](https://github.com/xuxueli/xxl-excel) 的实现原理，加入了一些自定义功能，并且对读取性能进行了优化的通用Excel读取、导出工具
> * 针对POI读取大文件时可能会引起内存溢出问题，更换读取方式为流式读取方式（Streaming Reader），采用缓冲流的方式进行读取；
> * ExcelReader采用流式方式进行读取，支持多种读取方式（文件路径、文件对象、输入流、以及指定数据起始结束行进行读取等）读取文件为集合；
> * ExcelWriter支持自定义导出Excel文件到用户桌面、指定路径、网页下载等；
> * 保留通过poi读取Excel的PoiUtils组件，对原生POI进行了封装。
#### 二、入门：
> ##### 1. 读取整个文档到集合中：
> * ``` List<Person> personList = ExcelReader.getListByFilePathAndClassType(filePath, Person.class);```
> ##### 2. 根据数据起始、结束行读取指定数据到集合中：
> * ``` List<Person> personList = ExcelReader.getListByFilePathAndClassType(filePath, Person.class, 1, 100);```
> ##### 3. 读取单列Excel为基本数据类型，性能最快：
> * ``` List<String> personList = ExcelReader.getListByFilePathAndSimpleClassType(filePath, String.class);```
#### 三、使用：
> ##### 1. 涉及到的注解:
> + @ExcelSheet：类注解：标注在待转换为Excel的Java类上
>   - `@ExcelSheet(name = "企业列表", headColor = HSSFColor.HSSFColorPredefined.LIGHT_GREEN)`
>   - name: 读取时，指定了名称，则读取指定名称的表单；导出时，指定了名称，则导出的文件名称为该名称；
>   - headColor: 导出时首行标题行的颜色；
> + @ExcelField：成员变量注解：标在待转换为Excel的Java类成员变量上
>   - `@ExcelField(name="名称", index=1, width=30*256, value="{'A':'待激活','B':'激活','C';'停机'}")`
>   - name: sql字段对应的中文名称；
>   - width: 列宽；
>   - index: 导出的时候可以通过指定index值来指定列之间的顺序；
>   - dateformat: 时间格式 yyyy-MM-dd或者其它格式；
>   - value: 自定义转换参数，这里定义成JSON字符串的格式，如数据库里定义的status字段(1:待审核 2:成功 3:失败)存储的是1、2、3等数字值，导出到列表时需要是汉字，可以在此转义；
>   - type: 导出时的标识字段，则只导出注解中包含"admin"标识的字段，如果不填，则导出所有；
> ##### 2. 定义Java类，标注上注解:
```
@ExcelSheet(name = "人员列表", headColor = HSSFColor.HSSFColorPredefined.LIGHT_GREEN)
   public class Person {
       @ExcelField(name = "编号", index = 1)
       private Integer id;
   
       @ExcelField(name = "姓名", index = 3, width = 30 * 256)
       private String name;
   
       // type = "admin",导出时，指定了type="admin", 则只导出包含此type = "admin"的列
       @ExcelField(name = "等级", index = 4, value = "{A:'普通会员',B:'白银会员',C:'黄金会员',D:'铂金会员',E:'钻石会员'}", width = 30 * 256, type = "admin")
       private String level;
   
       @ExcelField(name = "状态", index = 2, value = "{1:'正常',2:'禁用'}")
       private Integer status;
   
       // 不使用注解
       private String password;
   
       @ExcelField(name = "创建日期", index = 5, dateformat = "yyyy-MM-dd", type="admin")
       private Date createDate;
   
       // getter setter...
   }
```
> ##### 3. 具体使用，请参考上述入门中使用的方式
#### 四、方法列表：
> ##### 1. ExcelReader
方法名称|参数列表|方法说明
:----|:-----|:-----
左对齐|居中对齐|右对齐

