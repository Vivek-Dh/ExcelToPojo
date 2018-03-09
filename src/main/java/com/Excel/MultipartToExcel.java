package com.Excel;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

@RestController
public class MultipartToExcel {
	/**
	 * this method reads data from the uploaded excel file which is in the multipart
	 * file and creates a local excel file from it
	 * 
	 * @param Multipart
	 *            file
	 * @throws IOException
	 */
	@RequestMapping(value = "/file", method = RequestMethod.POST)
	public void saveExcel(@RequestParam("testDataFile") MultipartFile file) throws IOException {
		/*
		 * output location
		 */
		String outputFileLocation = "C:/songs/workbook.xlsx";
		XSSFWorkbook workbook = new XSSFWorkbook(file.getInputStream());
		FileOutputStream fileOut = new FileOutputStream(outputFileLocation);
		workbook.write(fileOut);
		fileOut.close();
	}
}
