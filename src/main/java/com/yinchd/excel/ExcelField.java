package com.yinchd.excel;

import java.lang.annotation.Inherited;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 字段注解，用于实体属性与Excel值之间的映射关系 </br>
 * eg：@ExcelField(name="名称", index=1, width=30*256, parseJson="{'1': '有效', '2': '无效')</br>
 * 其中name为必填，其它可以选填
 * @author yinchd
 * @since 2018-06-13
 */
@Target({java.lang.annotation.ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
@Inherited
public @interface ExcelField {

    /**
     * 导入时：name的值代表我们导入的excel文件中的列标题，在具体解析excel的过程中，会将实际读到的标题与ExcelField注解中name中定义值的作对比，据此判断导入文件是否合法；
     * 导出时：name的值代表是实体字段对应的中文名称，比如有个字段叫‘hobby’，ExcelField注解中name的值是‘爱好’，则导出的excel文件中hobby列的表头为‘爱好’；
     * eg：@ExcelField(name = "hobby")
     */
    String name();

    /**
     * 列宽，默认会自动根据内容适应宽度
     * eg：@ExcelField(width = 30*256)
     */
    int width() default 0;

    /**
     * 导出的时候用于排列字段（列）的顺序
     * eg：字段name的index为0，sex的index为1，则导出时生成的列字段name在前面，sex在后面
     */
    int index() default 0;

    /**
     * 时间格式
     * eg: "yyyy-MM-dd HH:mm:ss" , "yyyy-MM-dd"
     */
    String dateformat() default "yyyy-MM-dd HH:mm:ss";

    /**
     * 自定义转换参数，导出的时候比如字段值是枚举或者想转义成其它字符，这里可以定义一个json串，key为待转的字段，value为想转义出来值
     * eg：convertJson="{'1': '有效', '2': '无效', '3': '正常'}"
     */
    String parseJson() default "";

}