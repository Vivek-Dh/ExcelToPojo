package com.Excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.net.MalformedURLException;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

/**
 * @author vdehariya
 *
 */
@RestController
@RequestMapping(value = "/excel")
public class ExcelReader {
	/**
	 * This method maps the input data in the excel file to the kind of pojo
	 * specified
	 * 
	 * @param FILE_NAME
	 *            filepath/name of the excel file
	 * @param className
	 *            full name of the pojo class(package.classname)
	 * @return List of specified pojos
	 */
	@RequestMapping(value = "/conversion", method = RequestMethod.GET)
	public List<Object> excelToPojo(@RequestParam(value = "workbook", required = true) String FILE_NAME,
			@RequestParam(value = "class", required = true) String className) {
		try {
			FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
			XSSFWorkbook workbook = new XSSFWorkbook(excelFile);
			List<Object> pojoList = excelReaderService(workbook, className);
			return pojoList;
		} catch(MalformedURLException e) {
			e.printStackTrace();
		}
		catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} catch (NumberFormatException e) {
			e.printStackTrace();
		} catch (IllegalAccessException e) {
			e.printStackTrace();
		} catch (IllegalArgumentException e) {
			e.printStackTrace();
		} catch (InvocationTargetException e) {
			e.printStackTrace();
		} catch (ClassNotFoundException e) {
			e.printStackTrace();
		} catch (NoSuchMethodException e) {
			e.printStackTrace();
		} catch (SecurityException e) {
			e.printStackTrace();
		} catch (InstantiationException e) {
			e.printStackTrace();
		}
		return null;
	}

	/**
	 * This method parses the excel workbook and extracts the data fields and input
	 * data separately and calls the bean creation service.
	 * 
	 * @param workbook
	 * @param className
	 * @return List<Object> which contains the specified pojos
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
	public static List<Object> excelReaderService(XSSFWorkbook workbook, String className)
			throws NumberFormatException, IllegalAccessException, IllegalArgumentException, InvocationTargetException,
			ClassNotFoundException, NoSuchMethodException, SecurityException, InstantiationException, IOException {
		Sheet sheet = workbook.getSheetAt(0);
		Iterator<Row> rowIterator = sheet.iterator();
		Row row = null;
		/*
		 * pointing iterator to the first row to extract fields
		 */
		if (rowIterator.hasNext()) {
			row = rowIterator.next();
		}
		List<String> excelFields = new LinkedList<String>();
		Iterator<Cell> cellIterator = row.iterator();
		/*
		 * extracting field names from the excel sheet in the actual order
		 */
		while (cellIterator.hasNext()) {
			Cell cell = cellIterator.next();
			excelFields.add(cell.getStringCellValue());
		}
		/*
		 * Double dimension List to store the data values in the sheet
		 */
		List<LinkedList<String>> excelData = new LinkedList<LinkedList<String>>();
		/*
		 * extracting the data tuples from the excel sheet and changing the data type to
		 * string for ease in storage
		 */
		while (rowIterator.hasNext()) {
			LinkedList<String> list = new LinkedList<String>();
			row = rowIterator.next();
			cellIterator = row.iterator();
			while (cellIterator.hasNext()) {
				Cell cell = cellIterator.next();
				/*
				 * storing the data values as String
				 */
				if (cell.getCellTypeEnum() != CellType.STRING) {
					cell.setCellType(CellType.STRING);
				}
				list.add(cell.getStringCellValue());
			}
			excelData.add(list);
		}
		/*
		 * calling the bean creation method
		 */
		return ExcelService.createExcelBeans(excelData, excelFields, className);
	}
}
