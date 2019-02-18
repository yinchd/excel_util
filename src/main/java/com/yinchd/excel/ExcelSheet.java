package com.yinchd.excel;

import org.apache.poi.hssf.util.HSSFColor;

import java.lang.annotation.Inherited;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Excel与实体之间的映射 </br>
 * 添加注解的格式为 : </br>
 * eg：@ExcelSheet(name="用户列表", headColor=HSSFColor.HSSFColorPredefined.LIGHT_GREEN)</br>
 * name：Excel中sheet的名称，headColor：标题头的颜色
 * @author 研发部-殷晨东
 * @since 2018-06-13
 */
@Target({ java.lang.annotation.ElementType.TYPE })
@Retention(RetentionPolicy.RUNTIME)
@Inherited
public @interface ExcelSheet {
	/**
	 * ExcelSheet注解名称
	 */
	String name() default "";
	/**
	 * Title颜色
	 */
	HSSFColor.HSSFColorPredefined headColor() default HSSFColor.HSSFColorPredefined.LIGHT_GREEN;
}