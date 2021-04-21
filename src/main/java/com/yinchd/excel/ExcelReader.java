package com.yinchd.excel;

import com.monitorjbl.xlsx.StreamingReader;
import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Modifier;
import java.util.*;
import java.util.stream.Collectors;

/**
 * ExcelReader，通过流式方法读取excel，避免读取大文件时出现内存溢出问题
 * @author yinchd
 * @since 2018/2/14 14:26
 **/
@Slf4j
public class ExcelReader {

	/**
	 * 数据格式化工具类
	 */
	private static final DataFormatter DF = new DataFormatter();

	/**
	 * <p>
	 *	根据文件路径、数据起始值读取文件到集合 List<T>中
	 * 	eg: List<Person> personList = ExcelReader.getListByFilePath(filePath, Person.class); // 读全部行
	 * 	eg: List<Person> personList = ExcelReader.getListByFilePath(filePath, Person.class, 1, 100); // 1, 100 分别为行的下标
	 * </p>
	 * @param filePath 文件绝对路径 eg: e:/xxx/xxx.xlsx
	 * @param clazz 需要将表格内容读成的类型，如: Person.class
	 * @param dataStartOrEndIndex 数据起始行、数据结束行（非必填，注意，起始和结束行的值都是下标值，下标值是从0开始，如第一行的下标值为0）
	 * @param <T> 数据类型
	 * @return List<clazz>
	 */
	@SneakyThrows
	public static <T> List<T> getListByFilePath(String filePath, Class<T> clazz, int... dataStartOrEndIndex) {
		try (Workbook wb = getWorkbookByFile(new File(filePath))) {
			return readWorkbook(clazz, wb, dataStartOrEndIndex);
		}
	}

	/**
	 * <p>
	 * 	根据输入流、数据起始值读取文件到List<clazz>中
	 * 	eg: List<Person> personList = ExcelReader.getListByInputStream(inputStream, Person.class); // 读全部行
	 * 	eg: List<Person> personList = ExcelReader.getListByInputStream(inputStream, Person.class, 1, 100); // 1, 100 分别为行的下标
	 * </p>
	 * @param inputStream excel文件输入流
	 * @param clazz 需要将表格内容读成的类型，如：读成Person类型 eg: Person.class
	 * @param dataStartOrEndIndex 数据起始行、数据结束行（非必填，注意，起始和结束行的值都是下标值，下标值是从0开始，如第一行的下标值为0）
	 * @param <T> 数据类型
	 * @return List<clazz>
	 */
	@SneakyThrows
	public static <T> List<T> getListByInputStream(InputStream inputStream, Class<T> clazz, int... dataStartOrEndIndex) {
		try (Workbook wb = getWorkbookByInputStream(inputStream)) {
			return readWorkbook(clazz, wb, dataStartOrEndIndex);
		}
	}

	/**
	 * <p>
	 * 读取指定列的数据到List<String>中
	 * eg: List<String> personList = ExcelReader.getColumnListByFilePath(filePath, 3); // 读第一个sheet的第4列到集合中
	 * </p>
	 * @param filePath 文件绝对路径 eg: e:/xxx/xxx.xlsx
	 * @param columnIndex 要读取列的下标
	 * @param dataStartOrEndIndex 数据起始行、数据结束行（非必填，注意，起始和结束行的值都是下标值，下标值是从0开始，如第一行的下标值为0）
	 * @return 指定类型的集合
	 */
	@SneakyThrows
	public static List<String> getUniLineListByFilePath(String filePath, int columnIndex, int... dataStartOrEndIndex) {
		try (Workbook wb = getWorkbookByFile(new File(filePath))) {
			return getListByColumnIndex(wb, columnIndex, dataStartOrEndIndex);
		}
	}

	/**
	 * 读取指定列的数据到List<String>中
	 * eg: List<String> personList = ExcelReader.getColumnListByInputStream(inputStream, 3); // 读第一个sheet的第4列到集合中
	 * @param inputStream excel文件输入流
	 * @param columnIndex 要读取列的下标
	 * @param dataStartOrEndIndex 数据起始行、数据结束行（非必填，注意，起始和结束行的值都是下标值，下标值是从0开始，如第一行的下标值为0）
	 * @return 指定类型的集合
	 */
	@SneakyThrows
	public static List<String> getUniLineListByInputStream(InputStream inputStream, int columnIndex, int... dataStartOrEndIndex) {
		try (Workbook wb = getWorkbookByInputStream(inputStream)) {
			return getListByColumnIndex(wb, columnIndex, dataStartOrEndIndex);
		}
	}

	/**
	 * 根据输入流和指定要读取的sheet下标值来获取表格的物理数据总条数（可能包含空行）
	 * eg: int count = ExcelReader.getPhysicalRowCountByInputStream(inputStream); // 获取excel文件中第一个表单的物理数据总行数（有可能包含空行）
	 * @param inputStream excel文件输入流
	 * @return 可能包含空行的物理数据总条数
	 */
	public static int getPhysicalRowCountByInputStream(InputStream inputStream) {
		try (Workbook wb = getWorkbookByInputStream(inputStream)) {
			return getPhysicalRowCountByWorkbook(wb);
		} catch (Exception e) {
			throw new RuntimeException(e);
		}
	}

	/**
	 * 根据输入流获取表格有效数据总条数（不包含空行）
	 * eg: int count = ExcelReader.getRealRowCountByInputStream(inputStream); // 获取excel文件中第一个表单的真实数据总行数（不包含空行）
	 * @param inputStream excel文件输入流
	 * @return 不包含空行的有效数据总条数
	 */
	public static int getRealRowCountByInputStream(InputStream inputStream) {
		try (Workbook wb = getWorkbookByInputStream(inputStream)) {
			return getRealRowCountByWorkbook(wb);
		} catch (Exception e) {
			throw new RuntimeException(e);
		}
	}

	/**
	 * <p>
	 * 	通过输入流创建workbook，单独调用记得关闭流（上面通过try()的方式会自动关闭流，因为他们实现了AutoCloseble）
	 * </p>
	 * @param inputStream excel文件流
	 * @return Workbook对象
	 */
	private static Workbook getWorkbookByInputStream(InputStream inputStream) {
		try {
			return StreamingReader.builder()
					.rowCacheSize(100) // number of rows to keep in memory
					.bufferSize(4096) // buffer size to use when reading InputStream to file (defaults to 1024)
					.open(inputStream);
		} catch (Exception e) {
			throw new RuntimeException(e);
		}
	}

	/**
	 * <p>
	 * 	通过文件创建workbook，单独调用记得关闭流
	 * </p>
	 * @param file excel文件
	 * @return Workbook对象
	 */
	private static Workbook getWorkbookByFile(File file) {
		try (InputStream inputStream = new FileInputStream(file)) {
			return getWorkbookByInputStream(inputStream);
		} catch (Exception e) {
			throw new RuntimeException(e);
		}
	}

	/**
	 * 根据工作簿对象获取指定sheet表单中除去标题行后的总条数（可能包含空行），默认读取第一个sheet
	 * @param wb 工作簿对象
	 * @return 指定sheet中除去首行标题行后的物理数据条数（可能包含空行）
	 */
	private static int getPhysicalRowCountByWorkbook(Workbook wb) {
		Sheet sheet = wb.getSheetAt(0);
		int lastRowIndex = sheet.getLastRowNum();
		log.debug("排除s：{}", lastRowIndex);
		return lastRowIndex;
	}

	/**
	 * 获取指定sheet表单中除去空行的实际数据条数
	 * @param wb workbook
	 * @return 去空行后的数据总条数
	 */
	private static int getRealRowCountByWorkbook(Workbook wb) {
		Sheet sheet = wb.getSheetAt(0);
		int count = 0;
		for (Row row : sheet) {
			int j = 0, nullCellCount = 0;
			for (Iterator<Cell> ite = row.iterator(); ite.hasNext(); j++) {
				Cell cell = ite.next();
				if (cell == null || StringUtils.isBlank(DF.formatCellValue(cell))) {
					nullCellCount ++;
				}
			}
			if (nullCellCount >= j) {
				return count;
			}
			count++;
		}
		log.debug("excel sheet total row count excluded the empty row is：{}", count);
		return count;
	}

	/**
	 * 通过反射的方式将表格解析为List<clazz>集合，并支持指定起始与结束行读取
	 * 注意：Class<T> clazz 中必须要为需要映射的字段添加@ExcelField(name = "xxx")注解（xxx为excel文件中列的表头），逻辑中会将</br>
	 * clazz中被注解修饰的字段与excel文件中的列一一对应，name的值与excel文件中列的表头一一对应
	 * @param clazz 需要将表格内容读成的类型，如：读成Person类型 eg: Person.class
	 * @param wb 工作簿对象
	 * @param dataStartOrEndIndex 数据起始行、数据结束行（非必填，注意，起始和结束行的值都是下标值，下标值是从0开始，如第一行的下标值为0）
	 * @param <T> 数据类型
	 * @return 指定类型的集合
	 */
	@SneakyThrows
	private static <T> List<T> readWorkbook(Class<T> clazz, Workbook wb, int... dataStartOrEndIndex) {
		List<T> dataList = new ArrayList<>();
		Sheet sheet = getSheet(clazz, wb);
		// 取出所有被@ExcelField修饰的字段
		List<Field> fields = Arrays.stream(clazz.getDeclaredFields())
				.filter(field -> !Modifier.isStatic(field.getModifiers()) // 非static个包
						&& field.isAnnotationPresent(ExcelField.class) // 带@ExcelField修饰的字段
						&& field.getAnnotation(ExcelField.class).name().length() > 0) // @ExcelField中name的值不为空
				.sorted((f1, f2) -> { // 根据@ExcelField中指定的index进行排序，如果没有指定index的值，则按照字段读取的自然序来排列
					ExcelField f1Anno = f1.getAnnotation(ExcelField.class),
							f2Anno = f2.getAnnotation(ExcelField.class);
					return f1Anno.sortWeight() - f2Anno.sortWeight();
				}).collect(Collectors.toList());
		if (fields.size() == 0) {
			throw new RuntimeException("请检查实体类中有没有为待解析的字段添加@ExcelField注解");
		}
		// 存储标题行（默认第一行为标题行，且@ExcelField注释中name的值为excel文件中标题的值，后面会依此来校验导入的文件是不是符合我们自定义的要求
		List<String> titles = new ArrayList<>();
		// 存储name与field的映射关系，用于精确读取name列下的数据到field中
		Map<String, Field> fieldsMap = new HashMap<>();
		for (Field field : fields) {
			ExcelField anno = field.getAnnotation(ExcelField.class);
			fieldsMap.put(anno.name(), field);
			titles.add(anno.name().trim());
		}

		// 如果用户自定义了数据起始与结束行，则使用自定义值，如果没有指定，则默认从第二行开始读（默认第一行为标题行），到最后一行结束
		Integer[] idx = getStartAndEndIndex(sheet.getLastRowNum(), dataStartOrEndIndex);
		int dataStartIdx = idx[0], dataEndIdx = idx[1];
		T targetClass;
		int rowIndex = 0;
		// StreamReader读取的workbook只是借用了poi的类，不是所有的api都支持，所以这里用的迭代器的形式
		for (Iterator<Row> it = sheet.iterator(); it.hasNext() && rowIndex <= dataEndIdx; rowIndex++) {
			Row row = it.next();
			// 默认第一行为标题行，此处会拿到目标class里@ExcelField(name="id")的注解上的name值与此title里的数据校验
			if (rowIndex == 0) {
				int cellIndex = 0;
				for (Iterator<Cell> rowIt = row.iterator(); rowIt.hasNext() && cellIndex < titles.size(); cellIndex++) {
					Cell cell = rowIt.next();
					String cellValue = DF.formatCellValue(cell);
					if (!StringUtils.equals(cellValue, titles.get(cellIndex))) {
						throw new IllegalArgumentException("excel中第"+ (cellIndex + 1) +"列读取到的表头：["+ cellValue +"] " +
								"与您在@ExcelField(name = \"xxx\")注解上定义的值不一致，根据您定义的顺序："+ titles +"，该列的值应该为：" +
								"["+ titles.get(cellIndex) +"], 请检查您是否指定了注解中index的顺序，如未指定，则检查字段在类中定义的自然顺序");
					}
				}
			} else if (rowIndex >= dataStartIdx) {
				targetClass = clazz.newInstance();
				int j = 0, nullCellCount = 0;
				for (Iterator<Cell> ite = row.iterator(); ite.hasNext() && j < titles.size(); j++) {
					Cell cell = ite.next();
					String cellValue = DF.formatCellValue(cell);
					if (StringUtils.isBlank(cellValue)) {
						nullCellCount ++;
					}
					Field field = fieldsMap.get(titles.get(j));
					Object fieldValue = StringUtils.isBlank(cellValue) ? "" : FieldReflectionUtil.parseValue(field, cellValue.trim());
					field.setAccessible(true);
					field.set(targetClass, fieldValue);
				}
				if (nullCellCount >= j) {
					continue;
				}
				dataList.add(targetClass);
			}
		}
		return dataList;
	}

	/**
	 * 读取指定列的值到List中并过滤掉空值
	 * @param wb 工作簿对象
	 * @param dataStartOrEndIndex 数据起始行、数据结束行（非必填，注意，起始和结束行的值都是下标值，下标值是从0开始，如第一行的下标值为0）
	 * @return 字符串类型的集合
	 */
	private static List<String> getListByColumnIndex(Workbook wb, int columnIndex, int... dataStartOrEndIndex) {
		List<String> dataList = new ArrayList<>();
		// 默认读取第1个sheet，如果单独指定了需要读取的sheet名称，则读取自定义名称的sheet
		Sheet sheet = wb.getSheetAt(0);
		// 数据起始行，默认第一行为标题，第二行为数据起始点
		Integer[] idx = getStartAndEndIndex(sheet.getLastRowNum(), dataStartOrEndIndex);
		int dataStartIdx = idx[0], dataEndIdx = idx[1];
		int i = 0;
		for (Iterator<Row> it = sheet.iterator(); it.hasNext() && i <= dataEndIdx; i++) {
			Row row = it.next();
			// 验证标题行
			if (i >= dataStartIdx) {
				Cell cell = row.getCell(columnIndex);
				if (cell == null) {
					continue;
				}
				String cellValue = DF.formatCellValue(cell);
				if (StringUtils.isBlank(cellValue)) {
					continue;
				}
				dataList.add(cellValue);
			}
		}
		return dataList;
	}

	/**
	 * 计算读取数据的起始行与结束行，如果用户有传值，则读取用户自定义数据范围，如果用户没有传值，则使用默认值
	 * @param sheetLastRowIndex sheet中最后一行的下标值
	 * @param dataStartOrEndIndex 数据起始行、数据结束行（非必填，注意，起始和结束行的值都是下标值，下标值是从0开始，如第一行的下标值为0）
	 * @return 数据起始行与结束行
	 */
	private static Integer[] getStartAndEndIndex(int sheetLastRowIndex, int... dataStartOrEndIndex) {
		int startIndex = 1, endIndex = sheetLastRowIndex;
		if (dataStartOrEndIndex != null && dataStartOrEndIndex.length >= 1) {
			if (dataStartOrEndIndex.length == 2 && dataStartOrEndIndex[0] <= dataStartOrEndIndex[1]) {
				startIndex = dataStartOrEndIndex[0];
				endIndex = dataStartOrEndIndex[1];
			} else if (dataStartOrEndIndex.length == 1) {
				startIndex = dataStartOrEndIndex[0];
			}
		}
		return new Integer[]{startIndex, endIndex};
	}

	/**
	 * 读取工作簿中的默认sheet（默认读取第1个sheet）或者用户指定sheet名称的sheet
	 * @param clazz 目标类
	 * @param wb 工作簿
	 * @return sheet
	 */
	private static <T> Sheet getSheet(Class<T> clazz, Workbook wb) {
		Sheet sheet = null;
		if (clazz != null && clazz.isAnnotationPresent(ExcelSheet.class)) {
			ExcelSheet sheetAnno = clazz.getAnnotation(ExcelSheet.class);
			if (sheetAnno.name().length() > 0) {
				sheet = wb.getSheet(sheetAnno.name());
			}
		}
		if (sheet == null) {
			sheet = wb.getSheetAt(0);
		}
		return sheet;
	}

}
