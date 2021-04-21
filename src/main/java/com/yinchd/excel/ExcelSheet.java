package com.yinchd.excel;

import org.apache.poi.hssf.util.HSSFColor;

import java.lang.annotation.Inherited;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Excel与实体之间的映射 </br>
 * 添加注解的格式为: eg：@ExcelSheet(name = "用户列表", headColor = HSSFColor.HSSFColorPredefined.LIGHT_GREEN)</br>
 * @author yinchd
 * @since 2018/2/14
 */
@Target({java.lang.annotation.ElementType.TYPE})
@Retention(RetentionPolicy.RUNTIME)
@Inherited
public @interface ExcelSheet {

    /**
     * 对应excel中sheet表单的名称
     */
    String name() default "";

    /**
     * 统一的表头颜色
     */
    HSSFColor.HSSFColorPredefined headColor() default HSSFColor.HSSFColorPredefined.LIGHT_GREEN;

}