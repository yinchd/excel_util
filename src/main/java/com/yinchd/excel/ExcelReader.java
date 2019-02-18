package com.yinchd.excel;

import com.monitorjbl.xlsx.StreamingReader;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Modifier;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * <p>通过stream的形式读取excel，避免读取大文件时出现内存溢出问题</p>
 * @author 殷晨东
 * @since 2019/2/14 14:26
 **/
public class ExcelReader {

	private static Logger log = LoggerFactory.getLogger(ExcelReader.class);

	private static DataFormatter formator = new DataFormatter();

	/**
	 * 通过输入流创建workbook，单独调用记得关闭流
	 * @param inputStream excel文件流
	 * @return 工作簿对象
	 */
	private static Workbook getWorkbookByInputStream(InputStream inputStream) {
		try {
			Workbook workbook = StreamingReader.builder()
					.rowCacheSize(100) // number of rows to keep in memory
					.bufferSize(4096) // buffer size to use when reading InputStream to file (defaults to 1024)
					.open(inputStream);
			return workbook;
		} catch (Exception e) {
			throw new RuntimeException(e);
		}
	}

	/**
	 * 通过文件创建workbook，单独调用记得关闭流
	 * @param file excel文件
	 * @return 工作簿对象
	 */
	private static Workbook getWorkbookByFile(File file) {
		try (InputStream is = new FileInputStream(file)) {
			return getWorkbookByInputStream(is);
		} catch (Exception e) {
			throw new RuntimeException(e);
		}
	}

	/**
	 * 根据文件路径、数据起始值读取文件到List<clazz>中
	 * eg: List<Person> personList = ExcelReader.getListByFilePathAndClassType(filePath, Person.class, 1, 100);
	 * @param filePath 文件绝对路径 eg: e:/xxx/xxx.xlsx
	 * @param clazz 需要将表格内容读成的类型，如：读成Person类型 eg: Person.class
	 * @param dataStartOrEndIndex 数据起始行、数据结束行，数据起始行下标值，也即排除了标题行，数据起始值为0，此值可不传，不传时读取整个Excel数据到List中
	 * @param <T> 数据类型
	 * @return 指定类型的集合
	 */
	public static <T> List<T> getListByFilePathAndClassType(String filePath, Class<T> clazz, int... dataStartOrEndIndex) {
		try (Workbook wb = getWorkbookByFile(new File(filePath))) {
			return readWorkbook(clazz, wb, dataStartOrEndIndex);
		} catch (Exception e) {
			throw new RuntimeException(e);
		}
	}

	/**
	 * 根据输入流、数据起始值读取文件到List<clazz>中
	 * eg: List<Person> personList = ExcelReader.getListByInputStreamAndClassType(inputStream, Person.class, 1, 100);
	 * @param inputStream excel文件输入流
	 * @param clazz 需要将表格内容读成的类型，如：读成Person类型 eg: Person.class
	 * @param dataStartOrEndIndex 数据起始行、数据结束行，数据起始行下标值，也即排除了标题行，数据起始值为0，此值可不传，不传时读取整个Excel数据到List中
	 * @param <T> 数据类型
	 * @return 指定类型的集合
	 */
	public static <T> List<T> getListByInputStreamAndClassType(InputStream inputStream, Class<T> clazz, int... dataStartOrEndIndex) {
		try (Workbook wb = getWorkbookByInputStream(inputStream)) {
			return readWorkbook(clazz, wb, dataStartOrEndIndex);
		} catch (Exception e) {
			throw new RuntimeException(e);
		}
	}

	/**
	 * 根据文件路径、数据起始值读取文件到List<clazz>中，其中class为简单数据类型，可提升读取效率：如String，Integer，Date等，可自行扩展
	 * eg: List<String> personList = ExcelReader.getListByFilePathAndSimpleClassType(filePath, String.class);
	 * 此方法适合读取只有单列类型的简单表格，效率高
	 * @param filePath 文件绝对路径 eg: e:/xxx/xxx.xlsx
	 * @param clazz 需要将表格内容读成的类型，如：读成Person类型 eg: Person.class
	 * @param dataStartOrEndIndex 数据起始行、数据结束行，数据起始行下标值，也即排除了标题行，数据起始值为0，此值可不传，不传时读取整个Excel数据到List中
	 * @param <T> 数据类型
	 * @return 指定类型的集合
	 */
	public static <T> List<T> getListByFilePathAndSimpleClassType(String filePath, Class<T> clazz, int... dataStartOrEndIndex) {
		try (Workbook wb = getWorkbookByFile(new File(filePath))) {
			return getListByWorkbookAndSimpleClassType(wb, clazz, dataStartOrEndIndex);
		} catch (Exception e) {
			throw new RuntimeException(e);
		}
	}

	/**
	 * 根据输入流、数据起始值读取文件到List<clazz>中，其中class为简单数据类型，可提升读取效率：如String，Integer，Date等，可自行扩展
	 * eg: List<String> personList = ExcelReader.getListByInputStreamAndSimpleClassType(inputStream, String.class);
	 * 此方法适合读取只有单列类型的简单表格，效率高
	 * @param inputStream excel文件输入流
	 * @param clazz 需要将表格内容读成的类型，如：读成Person类型 eg: Person.class
	 * @param dataStartOrEndIndex 数据起始行、数据结束行，数据起始行下标值，也即排除了标题行，数据起始值为0，此值可不传，不传时读取整个Excel数据到List中
	 * @param <T> 数据类型
	 * @return 指定类型的集合
	 */
	public static <T> List<T> getListByInputStreamAndSimpleClassType(InputStream inputStream, Class<T> clazz, int... dataStartOrEndIndex) {
		try (Workbook wb = getWorkbookByInputStream(inputStream)) {
			return getListByWorkbookAndSimpleClassType(wb, clazz, dataStartOrEndIndex);
		} catch (Exception e) {
			throw new RuntimeException(e);
		}
	}

	/**
	 * 根据输入流获取表格的物理数据总条数（可能包含空行）
	 * eg: int excelPhysicalDataCount = ExcelReader.getPhysicalDataCountByInputStream(inputStream);
	 * @param inputStream excel文件输入流
	 * @return 可能包含空行的物理数据总条数
	 */
	public static int getPhysicalDataCountByInputStream(InputStream inputStream) {
		try (Workbook wb = getWorkbookByInputStream(inputStream)) {
			return getPhysicalDataCountByWorkbook(wb);
		} catch (Exception e) {
			throw new RuntimeException(e);
		}
	}

	/**
	 * 根据输入流获取表格有效数据总条数（不包含空行）
	 * @param inputStream excel文件输入流
	 * @return 不包含空行的有效数据总条数
	 */
	public static int getRealDataCountByInputStream(InputStream inputStream) {
		try (Workbook wb = getWorkbookByInputStream(inputStream)) {
			return getRealDataCountByWorkbook(wb);
		} catch (Exception e) {
			throw new RuntimeException(e);
		}
	}

	/**
	 * 根据工作簿对象获取第一个sheet表单中除去标题行后的总条数（可能包含空行）
	 * @param wb 工作簿对象
	 * @return 第一个sheet表单中除去首行标题行后的物理数据条数
	 */
	private static int getPhysicalDataCountByWorkbook(Workbook wb) {
		Sheet sheet = wb.getSheetAt(0);
		if (sheet == null) {
			log.debug("sheet获取为空...");
			throw new IllegalArgumentException("表格解析有误");
		}
		int lastRowIndex = sheet.getLastRowNum();
		log.debug("去除首行标题行后表格的总条数为：{}", lastRowIndex);
		return lastRowIndex;
	}

	/**
	 * 根据工作簿对象获取第一个sheet表单中除去标题行及空行的实际数据条数（不包含空行）
	 * @param wb workbook
	 * @return 第一个sheet表单中除去首行标题行后的数据条数
	 */
	private static int getRealDataCountByWorkbook(Workbook wb) {
		// 默认读取第1个sheet，如果单独指定了需要读取的sheet名称，则读取自定义名称的sheet
		Sheet sheet = wb.getSheetAt(0);
		if (sheet == null) {
			log.debug("sheet获取为空...");
			throw new IllegalArgumentException("表格解析有误");
		}
		int count = -1; // 去除标题行
		for (Iterator<Row> it = sheet.iterator(); it.hasNext();) {
			Row row = it.next();
			int j = 0, nullCellCount = 0;
			for (Iterator<Cell> ite = row.iterator(); ite.hasNext(); j++) {
				Cell cell = ite.next();
				if (cell == null || StringUtils.isBlank(formator.formatCellValue(cell))) {
					nullCellCount ++;
				}
			}
			if (nullCellCount >= j) {
				continue;
			}
			count ++;
		}
		log.debug("去除首行标题行以及空行后表格的总条数为：{}", count);
		return count;
	}

	/**
	 * 通过反射的方式将表格解析为List<T>集合，并支持指定起始与结束行读取
	 * @param clazz
	 * @param wb
	 * @param dataStartOrEndIndex
	 * @param <T>
	 * @return
	 */

	/**
	 * 通过反射的方式将表格解析为List<clazz>集合，并支持指定起始与结束行读取
	 * @param clazz 需要将表格内容读成的类型，如：读成Person类型 eg: Person.class
	 * @param wb 工作簿对象
	 * @param dataStartOrEndIndex 数据起始行、数据结束行，数据起始行下标值，也即排除了标题行，数据起始值为0，此值可不传，不传时读取整个Excel数据到List中
	 * @param <T> 数据类型
	 * @return 指定类型的集合
	 */
	public static <T> List<T> readWorkbook(Class<T> clazz, Workbook wb, int... dataStartOrEndIndex) {
		List<T> dataList = new ArrayList<>();
		try {
			// 默认读取第1个sheet，如果单独指定了需要读取的sheet名称，则读取自定义名称的sheet
			Sheet sheet = wb.getSheetAt(0);
			if (clazz.isAnnotationPresent(ExcelSheet.class)) {
				ExcelSheet excelSheet = clazz.getAnnotation(ExcelSheet.class);
				String sheetName = excelSheet == null ? "" : excelSheet.name();
				if (StringUtils.isNotBlank(sheetName)) {
					sheet = wb.getSheet(sheetName);
				}
			}
			if (sheet == null) {
				log.debug("sheet获取为空...");
				throw new IllegalArgumentException("表格解析有误");
			}
			Field[] fields = clazz.getDeclaredFields();
			List<Field> fieldList = new ArrayList<>();
			// title：用来校验Excel文件格式有没有错误
			List<String> title = new ArrayList<>();
			if (fields != null && fields.length > 0) {
				for (Field field :fields) {
					if (!Modifier.isStatic(field.getModifiers()) && field.isAnnotationPresent(ExcelField.class)) {
						fieldList.add(field);
						ExcelField anno = field.getAnnotation(ExcelField.class);
						String name = anno.name();
						if (StringUtils.isNotBlank(name)) {
							title.add(name.trim());
						}
					}
				}
			}

			if (fieldList == null || fieldList.size() <= 0)
				throw new IllegalArgumentException("@ExcelField注解字段为空，将导致无法解析");
			if (title.size() != fieldList.size())
				throw new IllegalArgumentException("待转换的class文件中缺少必要的注解描述");
			// 数据起始行，默认第一行为标题，第二行为数据起始点
			int dataStartIdx = 1, dataEndIdx = Integer.MAX_VALUE;
			if (dataStartOrEndIndex != null && dataStartOrEndIndex.length >= 1) {
				switch (dataStartOrEndIndex.length) {
					case 1:
						dataStartIdx = dataStartOrEndIndex[0];
						break;
					case 2:
						dataStartIdx = dataStartOrEndIndex[0];
						dataEndIdx = dataStartOrEndIndex[1];
						break;
					default:
						dataStartIdx = dataStartOrEndIndex[0];
				}
			}
			T targetClass; int i = 0;
			for (Iterator<Row> it = sheet.iterator(); it.hasNext() && i <= dataEndIdx; i++) {
				Row row = it.next();
				if (i == 0) {
					// 默认第一行为标题行，此处会拿到目标class里field的注解上的name值与此title里的数据校验
					int cellNum = 0;
					for (Iterator<Cell> rowIt = row.iterator(); rowIt.hasNext(); cellNum++) {
						Cell cell = rowIt.next();
						String cellValue = formator.formatCellValue(cell);
						if (!StringUtils.equals(cellValue,title.get(cellNum))) {
							throw new IllegalArgumentException("表格有误，请重新下载模板！");
						}
					}
					if (title.size() != cellNum) {
						throw new IllegalArgumentException("表格有误，请重新下载模板！");
					}
				} else if (i >= dataStartIdx) {
					targetClass = clazz.newInstance();
					int j = 0, nullCellCount = 0;
					for (Iterator<Cell> ite = row.iterator(); ite.hasNext(); j++) {
						Cell cell = ite.next();
						String cellValue = formator.formatCellValue(cell);
						if (StringUtils.isBlank(cellValue)) nullCellCount ++;
						Field field = fieldList.get(j);
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
		} catch (Exception e) {
			throw new RuntimeException(e);
		}
		return dataList;
	}

	/**
	 * 通过反射的方式将表格解析为List<T>集合，并支持指定起始与结束行读取
	 * @param wb 工作簿对象
	 * @param clazz 需要将表格内容读成的类型，如：读成Person类型 eg: Person.class
	 * @param dataStartOrEndIndex 数据起始行、数据结束行，数据起始行下标值，也即排除了标题行，数据起始值为0，此值可不传，不传时读取整个Excel数据到List中
	 * @param <T> 数据类型
	 * @return 指定类型的集合
	 */
	private static <T> List<T> getListByWorkbookAndSimpleClassType(Workbook wb, Class<T> clazz, int... dataStartOrEndIndex) {
		List<T> dataList = new ArrayList<>();
		// 默认读取第1个sheet，如果单独指定了需要读取的sheet名称，则读取自定义名称的sheet
		Sheet sheet = wb.getSheetAt(0);
		if (sheet == null) {
			log.debug("sheet获取为空...");
			throw new IllegalArgumentException("表格解析有误");
		}
		// 数据起始行，默认第一行为标题，第二行为数据起始点
		int dataStartIdx = 1, dataEndIdx = Integer.MAX_VALUE;
		if (dataStartOrEndIndex != null && dataStartOrEndIndex.length >= 1) {
			switch (dataStartOrEndIndex.length) {
				case 1:
					dataStartIdx = dataStartOrEndIndex[0];
					break;
				case 2:
					dataStartIdx = dataStartOrEndIndex[0];
					dataEndIdx = dataStartOrEndIndex[1];
					break;
				default:
					dataStartIdx = dataStartOrEndIndex[0];
			}
		}
		int i = 0; T t;
		for (Iterator<Row> it = sheet.iterator(); it.hasNext() && i <= dataEndIdx; i++) {
			Row row = it.next();
			// 验证标题行
			if (i >= dataStartIdx) {
				try {
					t = clazz.newInstance();
					// 只考虑读取一列值的情况
					Cell cell = row.getCell(0);
					String cellValue = formator.formatCellValue(cell);
					if (StringUtils.isBlank(cellValue)) continue;
					if (t instanceof String) {
						t = (T) cellValue;
					} else if (t instanceof Integer) {
						t = (T) Integer.valueOf(cellValue);
					}
					dataList.add(t);
				} catch (Exception e) {
					throw new RuntimeException(e);
				}
			}
		}
		return dataList;
	}

	/**
	 * 读取单列表格的标题行
	 * @param inputStream
	 * @return
	 */
	public static String getSingleColunmTitleByInputStream(InputStream inputStream) {
		List<String> list = getListByInputStreamAndSimpleClassType(inputStream, String.class, 0, 0);
		return list.size() == 1 ? list.get(0) : "";
	}

	/**
	 * 通过校验title的形式校验文件符不符合要求，当title对不上或者有空行时返回false
	 * @param title 正确的标题
	 * @param inputStream 文件流
	 * @return boolean
	 */
	public static boolean validExcelByTitle(String title, InputStream inputStream) {
		return StringUtils.equals(title, getSingleColunmTitleByInputStream(inputStream));
	}

}
