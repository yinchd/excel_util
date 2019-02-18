package com.yinchd.excel;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.BufferedOutputStream;
import java.io.FileOutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Modifier;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

/**
 * <p>Excel写出工具类</p>
 * @author 殷晨东
 * @since 2018/2/14 14:26
 */
public class ExcelWriter {

    private static Logger log = LoggerFactory.getLogger(ExcelWriter.class);

    /**
     * 导出类型
     */
    private static String EXPORT_ALL = null;

    /**
     * 导出Excel文件到用户桌面 C:\Users\xxx\Desktop
     * @param fileName 文件名称
     * @param dataList 数据集合
     * @param <T> 对象类型
     */
    public static <T> void writeToDesktop(String fileName, List<T> dataList) {
        if (!(fileName.endsWith(".xls") || fileName.endsWith("xlsx"))) {
            fileName += ".xlsx";
        }
        String filePath = System.getProperties().getProperty("user.home") + "\\Desktop\\" + fileName;
        try (HSSFWorkbook wb = createWorkbook(fileName, dataList, EXPORT_ALL);
             FileOutputStream fos = new FileOutputStream(filePath)) {
            wb.write(fos);
            fos.flush();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 根据导出类型导出Excel文件到用户桌面 C:\Users\xxx\Desktop
     * @param fileName 文件名称
     * @param dataList 数据集合
     * @param type 导出数据类型 type为空:代表不区别ExcelField中的type限制，只要包含@ExcelField注解都字段都导出,如果用户传入了其它值，则代表导出ExcelField注解中包含此type值的所有字段
     *  type为空:代表不区别ExcelField中的type限制，只要包含@ExcelField注解都字段都导出,如果用户传入了其它值，则代表导出ExcelField注解中包含此type值的所有字段
     * @param <T> void
     */
    public static <T> void writeToDesktop(String fileName, List<T> dataList, String type) {
        if (!(fileName.endsWith(".xls") || fileName.endsWith("xlsx")))
            fileName += ".xlsx";
        String filePath = System.getProperties().getProperty("user.home") + "\\Desktop\\" + fileName;
        try (HSSFWorkbook wb = createWorkbook(fileName, dataList, type);
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
        if (!(fileNameWithPath.endsWith(".xls") || fileNameWithPath.endsWith("xlsx"))) {
            fileNameWithPath += ".xlsx";
        }
        try (HSSFWorkbook wb = createWorkbook(fileNameWithPath, dataList, EXPORT_ALL);
             FileOutputStream fos = new FileOutputStream(fileNameWithPath)) {
            wb.write(fos);
            fos.flush();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 根据类型导出Excel文件到指定路径
     * @param fileNameWithPath 完整路径的文件名
     * @param dataList 数据集合
     * @param type 导出数据类型，type为空:代表不区别ExcelField中的type限制，只要包含@ExcelField注解都字段都导出,如果用户传入了其它值，则代表导出ExcelField注解中包含此type值的所有字段
     * @param <T> void
     */
    public static <T> void writeToPath(String fileNameWithPath, List<T> dataList, String type) {
        if (!(fileNameWithPath.endsWith(".xls") || fileNameWithPath.endsWith("xlsx"))) {
            fileNameWithPath += ".xlsx";
        }
        try (HSSFWorkbook wb = createWorkbook(fileNameWithPath, dataList, type);
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
        if (!(fileName.endsWith(".xls") || fileName.endsWith("xlsx"))) {
            fileName += ".xls";
        }
        response.reset();
        response.setHeader("Content-Type", "application/vnd.ms-excel");
        response.setHeader("Content-Disposition", "attachment; filename=\"" +
                new String(fileName.getBytes("gbk"), StandardCharsets.ISO_8859_1) + "\"");
        try (HSSFWorkbook wb = createWorkbook(fileName, dataList, EXPORT_ALL);
             ServletOutputStream out = response.getOutputStream();
             BufferedOutputStream bos = new BufferedOutputStream(out)) {
            wb.write(bos);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 根据下载类型区分下载
     * @param fileName 文件名称
     * @param dataList 数据集合
     * @param type type为空:代表不区别ExcelField中的type限制，只要包含@ExcelField注解都字段都导出,如果用户传入了其它值，则代表导出ExcelField注解中包含此type值的所有字段
     * @param response response对象
     * @param <T> void
     * @throws Exception 异常
     */
    public static <T> void writeToPage(String fileName, List<T> dataList, String type, HttpServletResponse response) throws Exception {
        if (!(fileName.endsWith(".xls") || fileName.endsWith("xlsx"))) {
            fileName = fileName + ".xlsx";
        }
        response.reset();
        response.setHeader("Content-Type", "application/vnd.ms-excel");
        response.setHeader("Content-Disposition", "attachment; filename=\"" +
                new String(fileName.getBytes("gbk"), StandardCharsets.ISO_8859_1) + "\"");
        try (HSSFWorkbook wb = createWorkbook(fileName, dataList, type);
             ServletOutputStream out = response.getOutputStream();
             BufferedOutputStream bos = new BufferedOutputStream(out)) {
            wb.write(bos);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * <p>组装工作簿</p>
     * @param fileName 文件名称
     * @param dataList 数据集合
     * @param type 导出字段类型
     * @param <T> dataList中参数类型
     * @return 工作簿对象
     * @throws Exception
     */
    public static <T> HSSFWorkbook createWorkbook(String fileName, List<T> dataList, String type) throws Exception {
        if (CollectionUtils.isEmpty(dataList)) {
            throw new IllegalArgumentException("待导出数据不能为空");
        }
        log.info("待导出数据总条数：{}", dataList.size());
        HSSFWorkbook wb = new HSSFWorkbook();
        Class<? extends Object> dataClass = dataList.get(0).getClass();
        String sheetName = dataClass.getSimpleName();
        HSSFColor.HSSFColorPredefined headColor = null;
        if (dataClass.isAnnotationPresent(ExcelSheet.class)) {
            ExcelSheet sheetAnno = dataClass.getAnnotation(ExcelSheet.class);
            if (StringUtils.isNotBlank(sheetAnno.name())) {
                sheetName = sheetAnno.name().trim();
            }
            headColor = sheetAnno.headColor();
        }
        HSSFSheet sheet = wb.createSheet(sheetName);
        Field[] fields = dataClass.getDeclaredFields();
        List<Field> fieldList = new ArrayList<>();
        if (fields != null && fields.length > 0) {
            for (Field field : fields) {
                // 非static修饰类，且加了字段注解的才执行导出
                if (!Modifier.isStatic(field.getModifiers()) && field.isAnnotationPresent(ExcelField.class)) {
                    if (StringUtils.isBlank(type)) {
                        fieldList.add(field);
                    } else {
                        ExcelField fieldAnno = field.getAnnotation(ExcelField.class);
                        String exportType = StringUtils.isNotBlank(fieldAnno.type()) ? fieldAnno.type().trim() : "";
                        if (StringUtils.isNotBlank(exportType) && exportType.contains(type)) {
                            fieldList.add(field);
                        }
                    }
                }
            }
        }

        if (CollectionUtils.isEmpty(fieldList)) {
            throw new IllegalArgumentException("待导出数据field不能为空");
        }

        // 对导出列进行排序
        Collections.sort(fieldList, (fld1, fld2) -> {
            ExcelField fld1Anno = fld1.getAnnotation(ExcelField.class),
                    fld2Anno = fld2.getAnnotation(ExcelField.class);
            if (fld1Anno.index() > fld2Anno.index()) {
                return 1;
            } else if (fld2Anno.index() > fld1Anno.index()) {
                return -1;
            }
            return 0;
        });

        HSSFCellStyle headStyle = getHeadStyle(wb, headColor);

        // 创建标题行
        HSSFRow titleRow = sheet.createRow(0);
        Field field;
        HSSFCell titleCell;
        ExcelField fieldAnno;
        titleRow.setHeight((short) 300); // 设置标题高度
        boolean userSetWidth = false;
        for (int i = 0, len = fieldList.size(); i < len; i++) {
            field = fieldList.get(i);
            fieldAnno = field.getAnnotation(ExcelField.class);
            String fieldName = "";
            if (fieldAnno != null)
                fieldName = StringUtils.isNotBlank(fieldAnno.name()) ? fieldAnno.name().trim() : field.getName();
            int fieldWidth = (fieldAnno != null) ? fieldAnno.width() : 30 * 256;
            titleCell = titleRow.createCell(i, CellType.STRING);
            titleCell.setCellStyle(headStyle);
            if (fieldWidth > 0) {
                sheet.setColumnWidth(i, fieldWidth);
                userSetWidth = true;
            }
            titleCell.setCellValue(fieldName);
        }

        for (int i = 0, len = dataList.size(); i < len; i++) {
            HSSFRow rowI = sheet.createRow(i + 1);
            Object data = dataList.get(i);
            HSSFCell cell;
            for (int j = 0, l = fieldList.size(); j < l; j++) {
                Field fld = fieldList.get(j);
                fld.setAccessible(true);
                Object fieldValue = fld.get(data);
                String fieldValueStr = FieldReflectionUtil.formatValue(fld, fieldValue);
                cell = rowI.createCell(j, CellType.STRING);
                if (StringUtils.isBlank(fieldValueStr)) {
                    fieldValueStr = "--";
                }
                cell.setCellValue(fieldValueStr);
            }
        }

        if (!userSetWidth) {
            for (int i = 0; i < fieldList.size(); i++) {
                sheet.autoSizeColumn((short) i);
            }
        }
        return wb;
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

}
