package com.yinchd.excel;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONObject;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.BufferedOutputStream;
import java.io.FileOutputStream;
import java.io.UnsupportedEncodingException;
import java.lang.reflect.Field;
import java.lang.reflect.Modifier;
import java.nio.charset.StandardCharsets;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.stream.Collectors;

/**
 * <p>Excel写出工具类</p>
 * @author yinchd
 * @since 2018/2/14
 */
@Slf4j
public class ExcelWriter {

    /**
     * 导出Excel文件到用户桌面 C:\Users\xxx\Desktop
     * @param fileName 文件名称
     * @param dataList 数据集合
     * @param <T> 对象类型
     */
    public static <T> void writeToDesktop(String fileName, List<T> dataList) {
        fileName = suffixCheck(fileName);
        String filePath = System.getProperties().getProperty("user.home") + "\\Desktop\\" + fileName;
        try (HSSFWorkbook wb = createWorkbook(fileName, dataList);
             FileOutputStream fos = new FileOutputStream(filePath)) {
            wb.write(fos);
            fos.flush();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 导出Excel文件到指定路径
     * @param fileNameWithPath 完整路径的文件名
     * @param dataList 数据集合
     * @param <T> void
     */
    public static <T> void writeToPath(String fileNameWithPath, List<T> dataList) {
        fileNameWithPath = suffixCheck(fileNameWithPath);
        try (HSSFWorkbook wb = createWorkbook(fileNameWithPath, dataList);
             FileOutputStream fos = new FileOutputStream(fileNameWithPath)) {
            wb.write(fos);
            fos.flush();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 通过流写出到文件，用于页面下载
     * @param fileName 文件名称
     * @param dataList 数据集合
     * @param response response对象
     * @param <T> void
     * @throws Exception 异常
     */
    public static <T> void writeToPage(String fileName, List<T> dataList, HttpServletResponse response) throws Exception {
        fileName = suffixCheck(fileName);
        setReponseHeader(fileName, response);
        try (HSSFWorkbook wb = createWorkbook(fileName, dataList);
             ServletOutputStream out = response.getOutputStream();
             BufferedOutputStream bos = new BufferedOutputStream(out)) {
            wb.write(bos);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * <p>
     *  组装工作簿
     * </p>
     * @param fileName 文件名称
     * @param dataList 数据集合
     * @param <T> dataList中参数类型
     * @return 工作簿对象
     * @throws Exception 异常
     */
    public static <T> HSSFWorkbook createWorkbook(String fileName, List<T> dataList) throws Exception {
        if (CollectionUtils.isEmpty(dataList)) {
            throw new IllegalArgumentException("待导出的数据不能为空");
        }
        // 创建工作簿
        HSSFWorkbook wb = new HSSFWorkbook();
        Class<?> dataClass = dataList.get(0).getClass();
        // 表单的名称默认指定为当前日期
        String sheetName = LocalDate.now().toString();
        // 表头颜色
        HSSFColor.HSSFColorPredefined headColor = null;
        if (dataClass.isAnnotationPresent(ExcelSheet.class)) {
            ExcelSheet sheetAnno = dataClass.getAnnotation(ExcelSheet.class);
            if (sheetAnno.name().length() > 0) {
                sheetName = sheetAnno.name().trim();
            }
            // 如果@ExcelSheet注解中指定了headColor，则用指定的颜色
            headColor = sheetAnno.headColor();
        }
        // 创建表单
        HSSFSheet sheet = wb.createSheet(sheetName);
        Field[] fields = dataClass.getDeclaredFields();
        if (fields.length == 0) {
            throw new RuntimeException("目标类中不包含任何字段，请检查");
        }
        List<Field> fieldList = Arrays.stream(fields)
                // 过滤掉static修饰的字段，并且字段上被@ExcelField注解修饰才选为导出字段
                .filter(field -> !Modifier.isStatic(field.getModifiers()) && field.isAnnotationPresent(ExcelField.class))
                .collect(Collectors.toList());
        // 如果没有标注解，则导出全部非静态字段
        if (CollectionUtils.isEmpty(fieldList)) {
            fieldList = Arrays.stream(fields)
                    .filter(field -> !Modifier.isStatic(field.getModifiers()))
                    .collect(Collectors.toList());
        }

        // 根据@ExcelField注解中的sortWeight值对导出列进行排序，从小到大排序，sortWeight值越小的字段越排在前面
        fieldList.sort((fld1, fld2) -> {
            ExcelField ano1 = fld1.getAnnotation(ExcelField.class),
                    ano2 = fld2.getAnnotation(ExcelField.class);
            if (ano1.sortWeight() > ano2.sortWeight()) {
                return 1;
            } else if (ano2.sortWeight() > ano1.sortWeight()) {
                return -1;
            }
            return 0;
        });

        // 创建表头的样式
        HSSFCellStyle headStyle = getHeadStyle(wb, headColor);

        // 创建标题行
        HSSFRow titleRow = sheet.createRow(0);
        // 设置标题高度（当然这里的值也可以在注解里加一个属性，能让用户自定义配置高度）
        titleRow.setHeight((short) 300);
        boolean autoColumnWidth = true;
        for (int i = 0, len = fieldList.size(); i < len; i++) {
            Field field = fieldList.get(i);
            HSSFCell cell = titleRow.createCell(i, CellType.STRING);
            ExcelField anno = field.getAnnotation(ExcelField.class);
            cell.setCellStyle(headStyle);
            cell.setCellValue(anno == null ? field.getName() : anno.name());
            if (anno != null && anno.width() > 0) {
                autoColumnWidth = false;
                sheet.setColumnWidth(i, anno.width());
            }
        }
        // 写出数据行
        for (int i = 0, len = dataList.size(); i < len; i++) {
            // 默认标题行是第一行，所以数据起始行为第二行
            HSSFRow row = sheet.createRow(i + 1);
            Object data = dataList.get(i);
            for (int j = 0, l = fieldList.size(); j < l; j++) {
                Field field = fieldList.get(j);
                field.setAccessible(true);
                HSSFCell cell = row.createCell(j, CellType.STRING);
                cell.setCellValue(formatValue(field, field.get(data)));
            }
        }
        // 自适合列宽，这里只能等内容全部填充到sheet后才知道自动去适配列宽，所以这里放到最后
        if (autoColumnWidth) {
            for (int i = 0, len = fieldList.size(); i < len; i++) {
                sheet.autoSizeColumn(i);
            }
        }
        return wb;
    }

    /**
     * 格式化成字符串（特别处理了date和parseJson中的字段）
     * @param field 字段
     * @param value 数据实体中的值
     * @return 适合写入到表格中的字符串值
     */
    public static String formatValue(Field field, Object value) {
        Class<?> fieldType = field.getType();
        if (value == null) {
            return "";
        }
        String result = String.valueOf(value);
        ExcelField anno = field.getAnnotation(ExcelField.class);
        String datePattern = "yyyy-MM-dd HH:mm:ss";
        if (anno != null) {
            datePattern = anno.dateFormat();
            // 如果@ExcelField注解中定义了parseJson的json转义字符串，则这里进行转义
            if (StringUtils.isNotBlank(anno.parseJson())) {
                JSONObject jso = JSON.parseObject(anno.parseJson());
                result = jso.getString(String.valueOf(value));
            }
        }
        // 日期单独转换
        if (Date.class.equals(fieldType)) {
            SimpleDateFormat dateFormat = new SimpleDateFormat(datePattern);
            result = dateFormat.format(value);
        }
        return result;
    }

    private static HSSFCellStyle getHeadStyle(HSSFWorkbook wb, HSSFColor.HSSFColorPredefined headColor) {
        // title样式
        HSSFCellStyle headStyle = wb.createCellStyle();
        if (headColor != null) {
            headStyle.setFillForegroundColor(headColor.getIndex());
            headStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            headStyle.setFillBackgroundColor(headColor.getIndex());
        }
        headStyle.setAlignment(HorizontalAlignment.CENTER); // 左右居中
        headStyle.setVerticalAlignment(VerticalAlignment.CENTER); // 上下居中
        headStyle.setBorderBottom(BorderStyle.THIN); // 下边框
        headStyle.setBorderLeft(BorderStyle.THIN);// 左边框
        headStyle.setBorderTop(BorderStyle.THIN);// 上边框
        headStyle.setBorderRight(BorderStyle.THIN);// 右边框
        // 创建样式字体
        HSSFFont font = wb.createFont();
        font.setFontName("宋体"); // 设置字体名称
        font.setBold(true); // 设置字体为粗体
        font.setFontHeightInPoints((short) 11); // 设置字体大小
        return headStyle;
    }


    private static void setReponseHeader(String fileName, HttpServletResponse response) throws UnsupportedEncodingException {
        response.reset();
        response.setHeader("Content-Type", "application/vnd.ms-excel");
        response.setHeader("Content-Disposition", "attachment; filename=\"" +
                new String(fileName.getBytes("gbk"), StandardCharsets.ISO_8859_1) + "\"");
    }

    private static String suffixCheck(String fileName) {
        if (!(fileName.endsWith(".xls") || fileName.endsWith("xlsx"))) {
            fileName += ".xlsx";
        }
        return fileName;
    }
}
