package com.controller.sharepoint.convert;
// https://gist.github.com/gitchandu/4491431

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XLS2HTML {

	public static final String[] FILE_TYPES = new String[] { "xls", "xlsx" };
	
	public static void convert(String inputPath, String outputPath) throws Exception {
	
		String NEW_LINE = "\n";
		String HTML_FILE_EXTENSION = ".html";
		String TEMP_FILE_EXTENSION = ".tmp";
		String HTML_SNNIPET_1 = "<!DOCTYPE html><html><head><title>";
		String HTML_SNNIPET_2 = "</title></head><body><table>";
		String HTML_SNNIPET_3 = "</table></body></html>";
		String HTML_TR_S = "<tr>";
		String HTML_TR_E = "</tr>";
		String HTML_TD_S = "<td>";
		String HTML_TD_E = "</td>";
	
		File xlsfile = new File(inputPath);
		BufferedWriter writer;
		Workbook workbook;
		String fileName = xlsfile.getName();
		if (fileName.toLowerCase().endsWith(XLS2HTML.FILE_TYPES[0])) {
			workbook = new HSSFWorkbook(new FileInputStream(xlsfile));
		} else {
			workbook = new XSSFWorkbook(new FileInputStream(xlsfile));
		}

		workbook.close();
		
		System.out.println(fileName);
		System.out.println(outputPath);
		String outputFolder = outputPath.substring(0,outputPath.lastIndexOf('\\'));
		
		File tempFile = File.createTempFile(fileName + '-', HTML_FILE_EXTENSION + TEMP_FILE_EXTENSION, new File(outputFolder));
		writer = new BufferedWriter(new FileWriter(tempFile));
		writer.write(HTML_SNNIPET_1);
		writer.write(fileName);
		writer.write(HTML_SNNIPET_2);
		Sheet sheet = workbook.getSheetAt(0);
		Iterator<Row> rows = sheet.rowIterator();
		Iterator<Cell> cells = null;
		while (rows.hasNext()) {
			Row row = rows.next();
			cells = row.cellIterator();
			writer.write(NEW_LINE);
			writer.write(HTML_TR_S);
			while (cells.hasNext()) {
				Cell cell = cells.next();
				writer.write(HTML_TD_S);
				writer.write(cell.toString());
				writer.write(HTML_TD_E);
			}
			writer.write(HTML_TR_E);
		}
		writer.write(NEW_LINE);
		writer.write(HTML_SNNIPET_3);
		writer.close();
		String fileName_truncated = fileName.substring(0,fileName.lastIndexOf('.'));
		System.out.println(outputFolder + '\\' + fileName_truncated + HTML_FILE_EXTENSION);
		File newFile = new File(outputFolder + '\\' + fileName_truncated + HTML_FILE_EXTENSION);
		tempFile.renameTo(newFile);
	}
}
