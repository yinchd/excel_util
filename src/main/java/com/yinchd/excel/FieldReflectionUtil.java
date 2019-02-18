package com.yinchd.excel;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONObject;
import org.apache.commons.lang3.StringUtils;

import java.lang.reflect.Field;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

public final class FieldReflectionUtil {
	public static Byte parseByte(String value) {
		try {
			value = value.replaceAll("　", "");
			return Byte.valueOf(value);
		} catch (NumberFormatException e) {
			throw new RuntimeException("parseByte but input illegal input=" + value, e);
		}
	}

	public static Boolean parseBoolean(String value) {
		value = value.replaceAll("　", "");
		if (Boolean.TRUE.toString().equalsIgnoreCase(value)) {
			return Boolean.TRUE;
		}
		if (Boolean.FALSE.toString().equalsIgnoreCase(value)) {
			return Boolean.FALSE;
		}
		throw new RuntimeException("parseBoolean but input illegal input=" + value);
	}

	public static Integer parseInt(String value) {
		try {
			value = value.replaceAll("　", "");
			return Integer.valueOf(value);
		} catch (NumberFormatException e) {
			throw new RuntimeException("parseInt but input illegal input=" + value, e);
		}
	}

	public static Short parseShort(String value) {
		try {
			value = value.replaceAll("　", "");
			return Short.valueOf(value);
		} catch (NumberFormatException e) {
			throw new RuntimeException("parseShort but input illegal input=" + value, e);
		}
	}

	public static Long parseLong(String value) {
		try {
			value = value.replaceAll("　", "");
			return Long.valueOf(value);
		} catch (NumberFormatException e) {
			throw new RuntimeException("parseLong but input illegal input=" + value, e);
		}
	}

	public static Float parseFloat(String value) {
		try {
			value = value.replaceAll("　", "");
			return Float.valueOf(value);
		} catch (NumberFormatException e) {
			throw new RuntimeException("parseFloat but input illegal input=" + value, e);
		}
	}

	public static Double parseDouble(String value) {
		try {
			value = value.replaceAll("　", "");
			return Double.valueOf(value);
		} catch (NumberFormatException e) {
			throw new RuntimeException("parseDouble but input illegal input=" + value, e);
		}
	}

	public static Date parseDate(String value, ExcelField excelField) throws ParseException {
			String datePattern = "yyyy-MM-dd HH:mm:ss";
			if (excelField != null) {
				datePattern = excelField.dateformat();
			}
			SimpleDateFormat dateFormat = new SimpleDateFormat(datePattern);
			return dateFormat.parse(value);
	}

	public static Object parseValue(Field field, String value) throws ParseException {
		Class<?> fieldType = field.getType();

		ExcelField excelField = (ExcelField) field.getAnnotation(ExcelField.class);
		if ((value == null) || (value.trim().length() == 0)) {
			return null;
		}
		value = value.trim();
		if ((Boolean.class.equals(fieldType)) || (Boolean.TYPE.equals(fieldType))) {
			return parseBoolean(value);
		}
		if (String.class.equals(fieldType)) {
			return value;
		}
		if ((Short.class.equals(fieldType)) || (Short.TYPE.equals(fieldType))) {
			return parseShort(value);
		}
		if ((Integer.class.equals(fieldType)) || (Integer.TYPE.equals(fieldType))) {
			return parseInt(value);
		}
		if ((Long.class.equals(fieldType)) || (Long.TYPE.equals(fieldType))) {
			return parseLong(value);
		}
		if ((Float.class.equals(fieldType)) || (Float.TYPE.equals(fieldType))) {
			return parseFloat(value);
		}
		if ((Double.class.equals(fieldType)) || (Double.TYPE.equals(fieldType))) {
			return parseDouble(value);
		}
		if (Date.class.equals(fieldType)) {
			return parseDate(value, excelField);
		}
		throw new RuntimeException("request illeagal type, type must be Integer not int Long not long etc, type=" + fieldType);
	}

	public static String formatValue(Field field, Object value) {
		Class<?> fieldType = field.getType();

		ExcelField excelField = (ExcelField) field.getAnnotation(ExcelField.class);
		if (value == null) {
			return null;
		}
		if (Date.class.equals(fieldType)) {
			String datePattern = "yyyy-MM-dd HH:mm:ss";
			if ((excelField != null) && (excelField.dateformat() != null)) {
				datePattern = excelField.dateformat();
			}
			SimpleDateFormat dateFormat = new SimpleDateFormat(datePattern);
			return dateFormat.format(value);
		} else {
			String str = String.valueOf(value);
			String jsonStr = excelField.value();
			if (StringUtils.isNotBlank(jsonStr)) {
				JSONObject jso = JSON.parseObject(jsonStr);
				str = jso.getString(str);
			}
			return str;
		}
	}
}
