package com.bucuoa.common.excel;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelImportUtils {

	@SuppressWarnings("resource")
	public List<Map<String, Object>> readExcel(String fileName, Map<Integer, String> meta, int headerNum,
			int sheetIndex) throws InvalidFormatException, IOException {
		FileInputStream is = new FileInputStream(fileName);
		Workbook wb = null;
		Sheet sheet = null; 
		if (fileName !=null && !fileName.equals("")) {
			throw new RuntimeException("导入文档为空!");
		} else if (fileName.toLowerCase().endsWith("xls")) {
			 wb = new HSSFWorkbook(is);
		} else if (fileName.toLowerCase().endsWith("xlsx")) {
			wb = new XSSFWorkbook(is);
		} else {
			throw new RuntimeException("文档格式不正确!");
		}
		if (wb.getNumberOfSheets() < sheetIndex) {
			throw new RuntimeException("文档中没有工作表!");
		}
		sheet = wb.getSheetAt(sheetIndex);

		int lastRowNum = sheet.getLastRowNum() + headerNum;

		List<Map<String, Object>> list = new ArrayList<Map<String, Object>>();
		for (int i = headerNum; i <= lastRowNum; i++) {
			Row row = sheet.getRow(i);
			Short lastCellNum = getRowLastNum(row);
			if (lastCellNum != null) {
				Map<String, Object> map = new HashMap<String, Object>();
				for (int k = 0; k < lastCellNum; k++) {
					Cell cell = row.getCell(k);
					Object cellValue = getCellValue(cell, Long.class);
					System.out.print("\t" + cellValue);

					String fieldname = meta.get(k);
					if (fieldname != null) {
						map.put(fieldname, cellValue);
						
					}
				}
				list.add(map);
			}

			System.out.println("");
		}
		return list;
	}

	private Short getRowLastNum(Row row) {
		try {
			short lastCellNum = row.getLastCellNum();
			return lastCellNum;
		} catch (Exception e) {
			// e.printStackTrace();
		}
		return null;
	}

	public Object getCellValue(Cell cell, Class type) {
		Object val = "";
		try {

			if (cell != null) {
				if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
					if (type == Integer.class) {
						Double numericCellValue = cell.getNumericCellValue();
						val = numericCellValue.intValue();
					} 	if (type == Long.class) {
						Double numericCellValue = cell.getNumericCellValue();
						val = numericCellValue.longValue();
					} else {
						val = cell.getNumericCellValue();
					}
				} else if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
					val = cell.getStringCellValue();
				} else if (cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
					val = cell.getCellFormula();
				} else if (cell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
					val = cell.getBooleanCellValue();
				} else if (cell.getCellType() == Cell.CELL_TYPE_ERROR) {
					val = cell.getErrorCellValue();
				}
			}
		} catch (Exception e) {
			return val;
		}
		return val;
	}
}
