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
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.stream.Collectors;

/**
 * <p>Excel写出工具类</p>
 * @author yinchd
 * @since 2018/2/14 14:26
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
            throw new IllegalArgumentException("export data can not be null or empty");
        }
        // 创建工作簿
        HSSFWorkbook wb = new HSSFWorkbook();
        Class<?> dataClass = dataList.get(0).getClass();
        String sheetName = dataClass.getSimpleName();
        HSSFColor.HSSFColorPredefined headColor = null;
        if (dataClass.isAnnotationPresent(ExcelSheet.class)) {
            ExcelSheet sheetAnno = dataClass.getAnnotation(ExcelSheet.class);
            if (sheetAnno.name().length() > 0) {
                sheetName = sheetAnno.name().trim();
            }
            headColor = sheetAnno.headColor();
        }
        HSSFSheet sheet = wb.createSheet(sheetName);
        Field[] fields = dataClass.getDeclaredFields();
        if (fields.length == 0) {
            throw new RuntimeException("target class do not contains any fields, please check and retry");
        }
        List<Field> fieldList = Arrays.stream(fields)
                .filter(field -> !Modifier.isStatic(field.getModifiers()) && field.isAnnotationPresent(ExcelField.class))
                .collect(Collectors.toList());
        // 如果没有标注解，则导出全部非静态字段
        if (CollectionUtils.isEmpty(fieldList)) {
            fieldList = Arrays.stream(fields)
                    .filter(field -> !Modifier.isStatic(field.getModifiers()))
                    .collect(Collectors.toList());
        }

        // 对导出列进行排序
        fieldList.sort((fld1, fld2) -> {
            ExcelField ano1 = fld1.getAnnotation(ExcelField.class),
                    ano2 = fld2.getAnnotation(ExcelField.class);
            if (ano1.index() > ano2.index()) {
                return 1;
            } else if (ano2.index() > ano1.index()) {
                return -1;
            }
            return 0;
        });

        HSSFCellStyle headStyle = getHeadStyle(wb, headColor);

        // 创建标题行
        HSSFRow titleRow = sheet.createRow(0);
        titleRow.setHeight((short) 300); // 设置标题高度
        for (int i = 0, len = fieldList.size(); i < len; i++) {
            Field field = fieldList.get(i);
            HSSFCell cell = titleRow.createCell(i, CellType.STRING);
            ExcelField anno = field.getAnnotation(ExcelField.class);
            cell.setCellStyle(headStyle);
            cell.setCellValue(anno == null ? field.getName() : anno.name());
            if (anno != null && anno.width() > 0) {
                sheet.setColumnWidth(i, anno.width());
            } else {
                sheet.autoSizeColumn(i);
            }
        }

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
        return wb;
    }

    public static String formatValue(Field field, Object value) {
        Class<?> fieldType = field.getType();
        String result = "";
        if (value == null) {
            return result;
        }
        ExcelField anno = field.getAnnotation(ExcelField.class);
        String datePattern = "yyyy-MM-dd HH:mm:ss";
        if (anno != null) {
            datePattern = anno.dateformat();
            if (StringUtils.isNotBlank(anno.parseJson())) {
                JSONObject jso = JSON.parseObject(anno.parseJson());
                result = jso.getString(String.valueOf(value));
            }
        }
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
