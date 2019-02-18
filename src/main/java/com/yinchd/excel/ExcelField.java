package com.yinchd.excel;

import java.lang.annotation.Inherited;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 字段注解，用于实体属性与Excel值之间的映射关系 </br>
 * 添加注解的格式为 : </br>
 * eg：@ExcelField(name="名称", index=1, width=30*256,value="{'A':'待激活','B':'激活','C';'停机'}")</br>
 * name：对应实体属性的中文名称，可以与页面列表头保持一致，index：列的顺序，用于工具在导出时能按此顺序执行导出， value：字符串形式的json数据，用于枚举类型的数据转换，width：单元格宽度
 * @author 研发部-殷晨东
 * @since 2018-06-13
 */
@Target({ java.lang.annotation.ElementType.FIELD })
@Retention(RetentionPolicy.RUNTIME)
@Inherited
public @interface ExcelField {
	/*
	 * 字段对应的中文名称
	 */
	String name() default "";
	/*
	 * 列宽  如： 30*256
	 */
	int width() default 0;
	/*
	 * 顺序（起始值可以从0开始，也可以从1开始，要保证数据有可比性）
	 */
	int index() default 0;
	/*
	 * 时间格式
	 */
	String dateformat() default "yyyy-MM-dd HH:mm:ss";
	/*
	 * 自定义转换参数 使用格式为JOSNObject格式
	 */
	String value() default "";
	/*
	 * 用于在导出时区分哪些字段适用于本次导出，如注解为‘admin’，则导出的时候只导出包含‘admin’的字段，如果不填，则导出所有
	 */
	String type() default "";
}