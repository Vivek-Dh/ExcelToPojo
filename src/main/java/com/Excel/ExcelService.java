package com.Excel;

import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.Set;


public class ExcelService {
	/**
	 * @param excelData
	 * @param excelFields
	 * @param className
	 * @return List<Object> list of the created pojos
	 * @throws NumberFormatException
	 * @throws IllegalAccessException
	 * @throws IllegalArgumentException
	 * @throws InvocationTargetException
	 * @throws ClassNotFoundException
	 * @throws NoSuchMethodException
	 * @throws SecurityException
	 * @throws InstantiationException
	 * @throws IOException
	 */
	public static List<Object> createExcelBeans(List<LinkedList<String>> excelData, List<String> excelFields,
			String className)
			throws NumberFormatException, IllegalAccessException, IllegalArgumentException, InvocationTargetException,
			ClassNotFoundException, NoSuchMethodException, SecurityException, InstantiationException, IOException {
		List<Object> pojoList = new LinkedList<Object>();
		/*
		 * getting the class instance of the specified name
		 */
		Class<?> pojoClass = Class.forName(className);
		Map<String, String> excelMap = getFieldsMapping(excelFields, className);
		Map<String, Method> methodMap = getMethodsMapping(className);
		/*
		 * iterating through the input excel data and assigning it to the pojo
		 */
		for (LinkedList<String> list : excelData) {
			/*
			 * creating an empty bean
			 */
			Object eb = pojoClass.newInstance();
			for (int i = 0; i < list.size(); i++) {
				/*
				 * invoking the setter methods for the specific pojo fields this is based upon
				 * the order of fields in the excel sheet
				 */
				Method m = methodMap.get(excelMap.get(excelFields.get(i)));
				m.invoke(eb, list.get(i));
			}
			/*
			 * adding the created pojo to the final output list
			 */
			pojoList.add(eb);
		}
		return pojoList;
	}

	/**
	 * this method maps the excel fields to pojo fields
	 * 
	 * @param excelFields
	 * @param className
	 * @return Map<String,String> mapping of excelFields to corresponding pojo
	 *         fields
	 * @throws ClassNotFoundException
	 */
	public static Map<String, String> getFieldsMapping(List<String> excelFields, String className)
			throws ClassNotFoundException {
		Map<String, String> fieldsMap = new LinkedHashMap<String, String>();
		/*
		 * getting the class instance from the pojo name
		 */
		Class<?> pojoClass = Class.forName(className);
		Field[] pojoFields = pojoClass.getDeclaredFields();
		Set<String> pojoFieldNames = new HashSet<String>();
		for (Field field : pojoFields) {
			pojoFieldNames.add(field.getName());
		}
		/*
		 * matching the field name in excel and pojo
		 */
		for (String s : excelFields) {
			if (pojoFieldNames.contains(s.toLowerCase())) {
				fieldsMap.put(s, s.toLowerCase());
			}
		}
		return fieldsMap;
	}

	/**
	 * @param className
	 * @return
	 * @throws ClassNotFoundException
	 */
	public static Map<String, Method> getMethodsMapping(String className) throws ClassNotFoundException {

		Class<?> pojoClass = Class.forName(className);
		/*
		 * fetching the pojo methods from the class instance
		 */
		Method pojoMethods[] = pojoClass.getDeclaredMethods();
		/*
		 * storing the method names in a list for further mapping
		 */
		List<String> pojoMethodsNames = new LinkedList<String>();
		for (Method method : pojoMethods) {
			pojoMethodsNames.add(method.getName());
		}
		/*
		 * getting the pojo fields
		 */
		Field[] pojoFields = ExcelBean.class.getDeclaredFields();
		Map<String, Method> methodsMap = new LinkedHashMap<String, Method>();
		/*
		 * mapping the pojo fields to the setter methods
		 */
		for (Field field : pojoFields) {
			String fieldName = field.getName();
			/*
			 * changing the first alphabet of pojo fields to Uppercase to match it with the
			 * general function naming convention e.g number -> Number(#getNumber)
			 */
			char c = (char) (fieldName.charAt(0) - 32);
			fieldName = c + fieldName.substring(1);
			/*
			 * getting the index of the setter method in the method array based upon method
			 * name
			 */
			int methodIndex = pojoMethodsNames.indexOf("set" + fieldName);
			/*
			 * mapping the pojo field to the setter method
			 */
			methodsMap.put(field.getName(), pojoMethods[methodIndex]);
		}
		return methodsMap;
	}
}
