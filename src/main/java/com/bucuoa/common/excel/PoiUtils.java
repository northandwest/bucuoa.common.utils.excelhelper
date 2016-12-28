package com.bucuoa.common.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.Reader;
import java.math.BigDecimal;
import java.sql.Clob;
import java.sql.SQLException;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;


public class PoiUtils {
	private static final Logger LOGGER = LoggerFactory.getLogger(PoiUtils.class);
	/* LONG */
	protected static final String LONG = "java.lang.Long";
	/* SHORT */
	protected static final String SHORT = "java.lang.Short";
	/* INT */
	protected static final String INT = "java.lang.Integer";
	/* STRING */
	protected static final String STRING = "java.lang.String";
	/* DATE */
	protected static final String DATE = "java.sql.Timestamp";
	/* BIG */
	protected static final String BIG = "java.math.BigDecimal";
	/* CLOB */
	protected static final String CLOB = "oracle.sql.CLOB";

	public static void main(String[] args) throws FileNotFoundException {

		String path = "C:/Program Files/feiq/AutoRecv Files/IP梳理记录表(1)(2).xlsx";
		File file = new File(path);
		InputStream inputStream = new FileInputStream(file);
		int count = getRecordsCountReadStream(inputStream, 1, false, 0);
		System.out.print(count);
	}

	/**
	 * 通过文件路径获取Excel读取行数
	 * 
	 * @param path
	 *            文件路径，只接受xls或xlsx结尾
	 * @param isHeader
	 *            是否表头
	 * @param headerCount
	 *            表头行数
	 * @return count 如果文件路径为空，返回0；
	 */
	public static int getRecordsCountReadPath(String path, boolean isHeader, int headerCount) {

		int count = 0;

		if (path == null) {
			return count;
		} else if (!path.endsWith("xls") && !path.endsWith("xlsx") && !path.endsWith("XLS") && !path.endsWith("XLSX")) {
			return count;
		}

		try {
			File file = new File(path);
			InputStream inputStream = new FileInputStream(file);
			Workbook hwb = null;
			if (path.endsWith("xls") || path.endsWith("XLS")) {
				hwb = new HSSFWorkbook(inputStream);
			} else if (path.endsWith("xlsx") || path.endsWith("XLSX")) {
				hwb = new XSSFWorkbook(inputStream);
			}

			if (null == hwb) {
				return count;
			}

			Sheet sheet = hwb.getSheetAt(0);// 暂定只取首页签
			int begin = sheet.getFirstRowNum();
			if (isHeader) {
				begin += headerCount;
			}
			int end = sheet.getLastRowNum();
			for (int i = begin; i <= end; i++) {
				if (null == sheet.getRow(i)) {
					continue;
				}
				count++;
			}

		} catch (FileNotFoundException e) {
			LOGGER.error("excel解析:", e);
			return 0;
		} catch (IOException e) {
			LOGGER.error("excel解析:", e);
			return 0;
		}
		return count;
	}

	/**
	 * 通过文件流获取Excel读取行数
	 * 
	 * @param path
	 *            文件路径
	 * @param type
	 *            类型，0为xls，1为xlsx；
	 * @param isHeader
	 *            是否表头
	 * @param headerCount
	 *            表头行数
	 * @return count 如果文件路径为空，返回0；
	 */
	public static int getRecordsCountReadStream(InputStream inputStream, int type, boolean isHeader, int headerCount) {

		int count = 0;
		if (type != 0 && type != 1) {
			return count;
		}

		try {
			Workbook hwb = null;
			if (type == 0) {
				hwb = new HSSFWorkbook(inputStream);
			} else if (type == 1) {
				hwb = new XSSFWorkbook(inputStream);
			}

			if (null == hwb) {
				return count;
			}

			Sheet sheet = hwb.getSheetAt(0);
			int begin = sheet.getFirstRowNum();
			if (isHeader) {
				begin += headerCount;
			}
			int end = sheet.getLastRowNum();
			for (int i = begin; i <= end; i++) {
				if (null == sheet.getRow(i)) {
					continue;
				}
				count++;
			}
		} catch (FileNotFoundException e) {
			LOGGER.error("excel解析:", e);
			return 0;
		} catch (IOException e) {
			LOGGER.error("excel解析:", e);
			return 0;
		}
		return count;
	}

	/**
	 * 通过文件流获取Excel读取
	 * 
	 * @param path
	 *            文件路径
	 * @param type
	 *            类型，0为xls，1为xlsx；
	 * @param isHeader
	 *            是否表头
	 * @param headerCount
	 *            表头行数
	 * @return poiList 如果文件路径为空，返回0；
	 */
	public static List<String[]> readRecordsInputStream(InputStream inputStream, int type, boolean isHeader,
			int headerCount) {
		List<String[]> poiList = new ArrayList<String[]>();
		if (type != 0 && type != 1) {
			return null;
		}
		if (type == 0) {
			poiList = readXLSRecords(inputStream, isHeader, headerCount);
		} else if (type == 1) {
			poiList = readXLSXRecords(inputStream, isHeader, headerCount);
		}
		return poiList;
	}

	/**
	 * 通过文件路径获取Excel读取
	 * 
	 * @param path
	 *            文件路径，只接受xls或xlsx结尾
	 * @param isHeader
	 *            是否表头
	 * @param headerCount
	 *            表头行数
	 * @return count 如果文件路径为空，返回0；
	 */
	public static List<String[]> readRecordsInputPath(String path, boolean isHeader, int headerCount) {
		List<String[]> poiList = new ArrayList<String[]>();
		if (path == null) {
			return null;
		} else if (!path.endsWith("xls") && !path.endsWith("xlsx") && !path.endsWith("XLS") && !path.endsWith("XLSX")) {
			return null;
		}
		File file = new File(path);
		try {
			InputStream inputStream = new FileInputStream(file);

			if (path.endsWith("xls") || path.endsWith("XLS")) {
				poiList = readXLSRecords(inputStream, isHeader, headerCount);
			} else if (path.endsWith("xlsx") || path.endsWith("XLSX")) {
				poiList = readXLSXRecords(inputStream, isHeader, headerCount);
			}
		} catch (Exception e) {
			LOGGER.error("excel解析:", e);
			return null;
		}
		return poiList;
	}

	/**
	 * 解析EXCEL2003文件流
	 * 如果一行记录的行中或行尾出现空格，POI工具类可能会跳过空格不做处理，所以默认第一行是表头，所有待解析的记录都以表头为准
	 * 
	 * @param inputStream
	 *            输入流
	 * @param isHeader
	 *            是否要跳过表头
	 * @param headerCount
	 *            表头占用行数
	 * @return 返回一个字符串数组List
	 */
	public static List<String[]> readXLSRecords(InputStream inputStream, boolean isHeader, int headerCount) {
		List<String[]> poiList = new ArrayList<String[]>();
		try {
			HSSFWorkbook wbs = new HSSFWorkbook(inputStream);
			HSSFSheet childSheet = wbs.getSheetAt(0);
			// 获取表头
			int begin = childSheet.getFirstRowNum();
			HSSFRow firstRow = childSheet.getRow(begin);
			int cellTotal = firstRow.getPhysicalNumberOfCells();
			// 是否跳过表头解析数据
			if (isHeader) {
				begin += headerCount;
			}
			// 逐行获取单元格数据
			for (int i = begin; i <= childSheet.getLastRowNum(); i++) {
				HSSFRow row = childSheet.getRow(i); // 一行的所有单元格格式都是常规的情况下，返回的row为null
				if (null != row) {
					String[] cells = new String[cellTotal];
					for (int k = 0; k < cellTotal; k++) {
						HSSFCell cell = row.getCell(k);
						cells[k] = getStringXLSCellValue(cell);
					}
					poiList.add(cells);
				}
			}
		} catch (Exception e) {
			LOGGER.error("excel解析:", e);
			return null;
		}
		return poiList;
	}

	/**
	 * 解析EXCEL2003文件流
	 * 如果一行记录的行中或行尾出现空格，POI工具类可能会跳过空格不做处理，所以默认第一行是表头，所有待解析的记录都以表头为准
	 * 该解析方法只适用于表头占用一行的情况
	 * 
	 * @param inputStream
	 *            输入流
	 * @param isHeader
	 *            是否要跳过表头
	 * @param headerCount
	 *            表头占用行数
	 * @param maxColNum
	 *            最大列数，适用于多表头
	 * @return 返回一个字符串数组List
	 */
	public static List<String[]> readXLSRecords(InputStream inputStream, boolean isHeader, int headerCount,
			int maxColNum) {
		List<String[]> poiList = new ArrayList<String[]>();
		try {
			HSSFWorkbook wbs = new HSSFWorkbook(inputStream);
			HSSFSheet childSheet = wbs.getSheetAt(0);
			// 获取表头
			int begin = childSheet.getFirstRowNum();
			// HSSFRow firstRow = childSheet.getRow(begin);
			// int cellTotal = firstRow.getPhysicalNumberOfCells();
			// 是否跳过表头解析数据
			if (isHeader) {
				begin += headerCount;
			}
			// 逐行获取单元格数据
			for (int i = begin; i <= childSheet.getLastRowNum(); i++) {
				HSSFRow row = childSheet.getRow(i); // 一行的所有单元格格式都是常规的情况下，返回的row为null
				String[] cells = new String[maxColNum]; // 空行对应空串数组
				for (int k = 0; k < maxColNum; k++) {
					HSSFCell cell = row == null ? null : row.getCell(k);
					cells[k] = getStringXLSCellValue(cell);
				}
				poiList.add(cells);
			}
		} catch (Exception e) {
			LOGGER.error("excel解析:", e);
			return null;
		}
		return poiList;
	}

	/**
	 * 解析EXCEL2007文件流
	 * 如果一行记录的行中或行尾出现空格，POI工具类可能会跳过空格不做处理，所以默认第一行是表头，所有待解析的记录都以表头为准
	 * 该处理方法中，表头对应都占用一行
	 * 
	 * @param inputStream
	 *            输入流
	 * @param isHeader
	 *            是否要跳过表头
	 * @param headerCount
	 *            表头占用行数
	 * @return 返回一个字符串数组List
	 */
	public static List<String[]> readXLSXRecords(InputStream inputStream, boolean isHeader, int headerCount) {
		List<String[]> poiList = new ArrayList<String[]>();
		try {
			XSSFWorkbook wbs = new XSSFWorkbook(inputStream);
			XSSFSheet childSheet = wbs.getSheetAt(0);
			// 获取表头
			int begin = childSheet.getFirstRowNum();
			XSSFRow firstRow = childSheet.getRow(begin);
			int cellTotal = firstRow.getPhysicalNumberOfCells();
			// 是否跳过表头解析数据
			if (isHeader) {
				begin += headerCount;
			}
			for (int i = begin; i <= childSheet.getLastRowNum(); i++) {
				XSSFRow row = childSheet.getRow(i); // 一行的所有单元格格式都是常规的情况下，返回的row为null
				if (null != row) {
					String[] cells = new String[cellTotal];
					for (int k = 0; k < cellTotal; k++) {
						XSSFCell cell = row.getCell(k);
						cells[k] = getStringXLSXCellValue(cell);
					}
					poiList.add(cells);
				}
			}
		} catch (Exception e) {
			LOGGER.error("excel解析:", e);
			return null;
		}
		return poiList;
	}

	/**
	 * 解析EXCEL2007文件流
	 * 如果一行记录的行中或行尾出现空格，POI工具类可能会跳过空格不做处理，所以默认第一行是表头，所有待解析的记录都以表头为准
	 * 该处理方法中，表头对应都占用一行
	 * 
	 * @param inputStream
	 *            输入流
	 * @param isHeader
	 *            是否要跳过表头
	 * @param headerCount
	 *            表头占用行数
	 * @param maxColNum
	 *            最大列数，适用于多表头的情况
	 * @return 返回一个字符串数组List
	 */
	public static List<String[]> readXLSXRecords(InputStream inputStream, boolean isHeader, int headerCount,
			int maxColNum) {
		List<String[]> poiList = new ArrayList<String[]>();
		try {
			XSSFWorkbook wbs = new XSSFWorkbook(inputStream);
			XSSFSheet childSheet = wbs.getSheetAt(0);
			// 获取表头
			int begin = childSheet.getFirstRowNum();
			// XSSFRow firstRow = childSheet.getRow(begin);
			// int cellTotal = firstRow.getPhysicalNumberOfCells();
			// 是否跳过表头解析数据
			if (isHeader) {
				begin += headerCount;
			}
			for (int i = begin; i <= childSheet.getLastRowNum(); i++) {
				XSSFRow row = childSheet.getRow(i); // 一行的所有单元格格式都是常规的情况下，返回的row为null
				String[] cells = new String[maxColNum]; // 空行对应空串数组
				for (int k = 0; k < maxColNum; k++) {
					XSSFCell cell = row == null ? null : row.getCell(k);
					cells[k] = getStringXLSXCellValue(cell);
				}
				poiList.add(cells);
			}
		} catch (Exception e) {
			LOGGER.error("excel解析:", e);
			return null;
		}
		return poiList;
	}

	/**
	 * 获取单元格数据内容为字符串类型的数据
	 * 
	 * @param cell
	 *            Excel单元格
	 * @return String 单元格数据内容
	 */
	private static String getStringXLSCellValue(HSSFCell cell) {
		String strCell = "";
		if (cell == null) {
			return "";
		}

		// 将数值型参数转成文本格式，该算法不能保证1.00这种类型数值的精确度
		DecimalFormat df = (DecimalFormat) NumberFormat.getPercentInstance();
		StringBuffer sb = new StringBuffer();
		sb.append("0");
		df.applyPattern(sb.toString());

		switch (cell.getCellType()) {
		case HSSFCell.CELL_TYPE_STRING:
			strCell = cell.getStringCellValue();
			break;
		case HSSFCell.CELL_TYPE_NUMERIC:
			double value = cell.getNumericCellValue();
			while (Double.parseDouble(df.format(value)) != value) {
				if ("0".equals(sb.toString())) {
					sb.append(".0");
				} else {
					sb.append("0");
				}
				df.applyPattern(sb.toString());
			}
			strCell = df.format(value);
			break;
		case HSSFCell.CELL_TYPE_BOOLEAN:
			strCell = String.valueOf(cell.getBooleanCellValue());
			break;
		case HSSFCell.CELL_TYPE_BLANK:
			strCell = "";
			break;
		default:
			strCell = "";
			break;
		}
		if (strCell == null || "".equals(strCell)) {
			return "";
		}
		return strCell;
	}

	/**
	 * 获取单元格数据内容为字符串类型的数据
	 * 
	 * @param cell
	 *            Excel单元格
	 * @return String 单元格数据内容
	 */
	private static String getStringXLSXCellValue(XSSFCell cell) {
		String strCell = "";
		if (cell == null) {
			return "";
		}
		// 将数值型参数转成文本格式，该算法不能保证1.00这种类型数值的精确度
		DecimalFormat df = (DecimalFormat) NumberFormat.getPercentInstance();
		StringBuffer sb = new StringBuffer();
		sb.append("0");
		df.applyPattern(sb.toString());

		switch (cell.getCellType()) {
		case XSSFCell.CELL_TYPE_STRING:
			strCell = cell.getStringCellValue();
			break;
		case XSSFCell.CELL_TYPE_NUMERIC:
			double value = cell.getNumericCellValue();
			while (Double.parseDouble(df.format(value)) != value) {
				if ("0".equals(sb.toString())) {
					sb.append(".0");
				} else {
					sb.append("0");
				}
				df.applyPattern(sb.toString());
			}
			strCell = df.format(value);
			break;
		case XSSFCell.CELL_TYPE_BOOLEAN:
			strCell = String.valueOf(cell.getBooleanCellValue());
			break;
		case XSSFCell.CELL_TYPE_BLANK:
			strCell = "";
			break;
		default:
			strCell = "";
			break;
		}
		if (strCell == null || "".equals(strCell)) {
			return "";
		}
		return strCell;
	}

	/**
	 * 导出Excel
	 * 
	 * @param response
	 * @param request
	 * @param title
	 * @param map
	 *            key为标题，list为数据，单表头导出
	 * @param type
	 *            0为xls，1为xlsx
	 */
	public static void output(HttpServletResponse response, HttpServletRequest request, String title,
			Map<String, List<Object>> map, int type) throws Exception {

		if (type != 0 && type != 1) {
			throw new Exception("无效的excel导出类型，type=0表示xls，type=1表示xlsx");
		}
		List<Map<String, List<Object>>> list = new ArrayList<Map<String, List<Object>>>();
		list.add(map);
		if (type == 0) {
			outputXLS(response, request, title, list);
		} else if (type == 1) {
			outputXLSX(response, request, title, list);
		}
	}

	/**
	 * 导出Excel
	 * 
	 * @param response
	 * @param request
	 * @param title
	 * @param List<Map>
	 *            支持多表头导出
	 * @param type
	 *            0为xls，1为xlsx
	 */
	public static void output(HttpServletResponse response, HttpServletRequest request, String title,
			List<Map<String, List<Object>>> list, int type) throws Exception {

		if (type != 0 && type != 1) {
			throw new Exception("无效的excel导出类型，type=0表示xls，type=1表示xlsx");
		}
		if (type == 0) {
			outputXLS(response, request, title, list);
		} else if (type == 1) {
			outputXLSX(response, request, title, list);
		}
	}

	/**
	 * 导出Excel
	 * 
	 * @param response
	 * @param request
	 * @param title
	 * @param List<Map>
	 *            支持多表头按顺序导出
	 * @param type
	 *            0为xls，1为xlsx
	 */
	public static void outputByColId(HttpServletResponse response, HttpServletRequest request, String title,
			List<Map<String, List<Object>>> list, int type) throws Exception {

		if (type != 0 && type != 1) {
			throw new Exception("无效的excel导出类型，type=0表示xls，type=1表示xlsx");
		}
		if (type == 0) {
			outputXLSByCol(response, request, title, list);
		} else if (type == 1) {
			outputXLSByCol(response, request, title, list);
		}
	}

	public static void outputXLS(HttpServletResponse response, HttpServletRequest request, String title,
			List<Map<String, List<Object>>> list) throws Exception {

		// 输出流定义
		OutputStream os = response.getOutputStream();
		byte[] fileNameByte = (title + ".xls").getBytes("GBK");
		String filename = new String(fileNameByte, "ISO8859-1");
		response.setContentType("application/x-msdownload");
		response.setCharacterEncoding("UTF-8");
		response.setHeader("Content-Disposition", "attachment;filename=" + filename);

		// 创建excel文件
		HSSFWorkbook hssf_w_book = new HSSFWorkbook();
		HSSFSheet hssf_w_sheet = hssf_w_book.createSheet(title);
		hssf_w_sheet.setDefaultColumnWidth(21); // 固定列宽度
		HSSFRow hssf_w_row = null;// 创建一行
		HSSFCell hssf_w_cell = null;// 创建每个单元格

		// 定义表头单元格样式
		HSSFCellStyle head_cellStyle = hssf_w_book.createCellStyle();
		// 定义表头字体样式
		HSSFFont head_font = hssf_w_book.createFont();
		head_font.setFontName("宋体");// 设置头部字体为宋体
		head_font.setBoldweight(Font.BOLDWEIGHT_BOLD); // 粗体
		head_font.setFontHeightInPoints((short) 10); // 字体大小
		// 表头单元格样式设置
		head_cellStyle.setFont(head_font);// 单元格样式使用字体
		head_cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		head_cellStyle.setVerticalAlignment(HSSFCellStyle.ALIGN_CENTER);
		// head_cellStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		// head_cellStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		// head_cellStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);
		// head_cellStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);

		// 定义数据单元格样式
		HSSFCellStyle cellStyle_CN = hssf_w_book.createCellStyle();// 创建数据单元格样式(数据库数据样式)
		// cellStyle_CN.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		// cellStyle_CN.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		// cellStyle_CN.setBorderRight(XSSFCellStyle.BORDER_THIN);
		// cellStyle_CN.setBorderTop(XSSFCellStyle.BORDER_THIN);

		// 在多表头导出时，定义第一个表头出现位置
		int titleFlag = 0;

		// 遍历写入表数据的list
		for (Map<String, List<Object>> map : list) {
			// 遍历map获取表头字段，并将表头字段放进String型的数组
			Set<String> key = map.keySet();
			String titles = "";
			int count = 0;
			for (Iterator<String> it = key.iterator(); it.hasNext();) {
				if (count != 0) {
					titles += ";";
				}
				titles += (String) it.next();
				count++;
			}
			String[] titleArray = titles.split(";");

			// 表头写入位置
			hssf_w_row = hssf_w_sheet.createRow(titleFlag);
			for (int i = 0; i < titleArray.length; i++) {
				hssf_w_cell = hssf_w_row.createCell(i);
				hssf_w_cell.setCellType(XSSFCell.CELL_TYPE_STRING);
				hssf_w_cell.setCellValue(titleArray[i]);
				hssf_w_cell.setCellStyle(head_cellStyle);
				// hssf_w_sheet.autoSizeColumn(( short ) i );
			}

			// 循环写入表数据，获取表的总列数，然后逐行写入数据
			for (int i = 0; i < map.get(titleArray[0]).size(); i++) {
				// 定义数据行
				hssf_w_row = hssf_w_sheet.createRow(i + titleFlag + 1);
				// 按行将每一列的数据写入单元格
				for (int j = 0; j < titleArray.length; j++) {
					hssf_w_cell = hssf_w_row.createCell(j);
					Object in = map.get(titleArray[j]).get(i);
					type4ExcelXLS(in, hssf_w_cell, cellStyle_CN);
					// hssf_w_sheet.autoSizeColumn(( short ) i );
				}
			}
			// 下一个表头的写入位置，和上一个表头数据之间隔一行
			titleFlag += map.get(titleArray[0]).size() + 2;
		}

		// excel文件导出
		hssf_w_book.write(os);
		os.close();
		request.getSession().setAttribute("EXCEL_FINISH", "1");
	}

	public static void outputXLSByCol(HttpServletResponse response, HttpServletRequest request, String title,
			List<Map<String, List<Object>>> list) throws Exception {

		// 输出流定义
		OutputStream os = response.getOutputStream();
		byte[] fileNameByte = (title + ".xls").getBytes("GBK");
		String filename = new String(fileNameByte, "ISO8859-1");
		response.setContentType("application/x-msdownload");
		response.setCharacterEncoding("UTF-8");
		response.setHeader("Content-Disposition", "attachment;filename=" + filename);

		// 创建excel文件
		HSSFWorkbook hssf_w_book = new HSSFWorkbook();
		HSSFSheet hssf_w_sheet = hssf_w_book.createSheet(title);
		hssf_w_sheet.setDefaultColumnWidth(21); // 固定列宽度
		HSSFRow hssf_w_row = null;// 创建一行
		HSSFCell hssf_w_cell = null;// 创建每个单元格

		// 定义表头单元格样式
		HSSFCellStyle head_cellStyle = hssf_w_book.createCellStyle();
		// 定义表头字体样式
		HSSFFont head_font = hssf_w_book.createFont();
		head_font.setFontName("宋体");// 设置头部字体为宋体
		head_font.setBoldweight(Font.BOLDWEIGHT_BOLD); // 粗体
		head_font.setFontHeightInPoints((short) 10); // 字体大小
		// 表头单元格样式设置
		head_cellStyle.setFont(head_font);// 单元格样式使用字体
		head_cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		head_cellStyle.setVerticalAlignment(HSSFCellStyle.ALIGN_CENTER);
		// head_cellStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		// head_cellStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		// head_cellStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);
		// head_cellStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);

		// 定义数据单元格样式
		HSSFCellStyle cellStyle_CN = hssf_w_book.createCellStyle();// 创建数据单元格样式(数据库数据样式)
		// cellStyle_CN.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		// cellStyle_CN.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		// cellStyle_CN.setBorderRight(XSSFCellStyle.BORDER_THIN);
		// cellStyle_CN.setBorderTop(XSSFCellStyle.BORDER_THIN);

		// 在多表头导出时，定义第一个表头出现位置
		int titleFlag = 0;

		// 遍历写入表数据的list
		for (Map<String, List<Object>> map : list) {
			// 遍历map获取表头字段，并将表头字段放进String型的数组
			Set<String> key = map.keySet();
			String titles = "";
			int count = 0;
			for (Iterator<String> it = key.iterator(); it.hasNext();) {
				if (count != 0) {
					titles += ";";
				}
				titles += (String) it.next();
				count++;
			}
			String[] titleArray = titles.split(";");
			String[] temArr = new String[titleArray.length];
			String[] temCol = new String[titleArray.length];
			for (int k = 0; k < titleArray.length; k++) {
				String tem = titleArray[k];
				String[] t = tem.split("_");
				String c = t[1];
				int n = Integer.parseInt(c);
				temArr[n] = t[0] + "_" + c;
				temCol[n] = t[0];
			}

			// 表头写入位置
			hssf_w_row = hssf_w_sheet.createRow(titleFlag);
			for (int i = 0; i < temArr.length; i++) {
				hssf_w_cell = hssf_w_row.createCell(i);
				hssf_w_cell.setCellType(XSSFCell.CELL_TYPE_STRING);
				hssf_w_cell.setCellValue(temCol[i]);
				hssf_w_cell.setCellStyle(head_cellStyle);
				// hssf_w_sheet.autoSizeColumn(( short ) i );
			}

			// 循环写入表数据，获取表的总列数，然后逐行写入数据
			for (int i = 0; i < map.get(temArr[0]).size(); i++) {
				// 定义数据行
				hssf_w_row = hssf_w_sheet.createRow(i + titleFlag + 1);
				// 按行将每一列的数据写入单元格
				for (int j = 0; j < temArr.length; j++) {
					hssf_w_cell = hssf_w_row.createCell(j);
					Object in = map.get(temArr[j]).get(i);
					type4ExcelXLS(in, hssf_w_cell, cellStyle_CN);
					// hssf_w_sheet.autoSizeColumn(( short ) i );
				}
			}
			// 下一个表头的写入位置，和上一个表头数据之间隔一行
			titleFlag += map.get(temArr[0]).size() + 2;
		}

		// excel文件导出
		hssf_w_book.write(os);
		os.close();
		request.getSession().setAttribute("EXCEL_FINISH", "1");
	}

	public static void outputXLSX(HttpServletResponse response, HttpServletRequest request, String title,
			List<Map<String, List<Object>>> list) throws Exception {

		// 输出流定义
		OutputStream os = response.getOutputStream();
		byte[] fileNameByte = (title + ".xlsx").getBytes("GBK");
		String filename = new String(fileNameByte, "ISO8859-1");
		response.setContentType("application/x-msdownload");
		response.setCharacterEncoding("UTF-8");
		response.setHeader("Content-Disposition", "attachment;filename=" + filename);

		// 创建excel文件
		XSSFWorkbook xssf_w_book = new XSSFWorkbook();
		XSSFSheet xssf_w_sheet = xssf_w_book.createSheet(title);
		xssf_w_sheet.setDefaultColumnWidth(21); // 固定列宽度
		XSSFRow xssf_w_row = null;// 创建一行
		XSSFCell xssf_w_cell = null;// 创建每个单元格

		// 定义表头单元格样式
		XSSFCellStyle head_cellStyle = xssf_w_book.createCellStyle();
		// 定义表头字体样式
		XSSFFont head_font = xssf_w_book.createFont();
		head_font.setFontName("宋体");// 设置头部字体为宋体
		head_font.setBoldweight(Font.BOLDWEIGHT_BOLD); // 粗体
		head_font.setFontHeightInPoints((short) 10);
		// 表头单元格样式设置
		head_cellStyle.setFont(head_font);// 单元格使用表头字体样式
		head_cellStyle.setAlignment(HorizontalAlignment.CENTER);
		head_cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		head_cellStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		head_cellStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		head_cellStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);
		head_cellStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);

		// 定义数据单元格样式
		XSSFCellStyle cellStyle_CN = xssf_w_book.createCellStyle();// 创建数据单元格样式(数据库数据样式)
		cellStyle_CN.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		cellStyle_CN.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		cellStyle_CN.setBorderRight(XSSFCellStyle.BORDER_THIN);
		cellStyle_CN.setBorderTop(XSSFCellStyle.BORDER_THIN);

		// 在多表头导出时，定义第一个表头出现位置
		int titleFlag = 0;

		// 遍历写入表数据的list
		for (Map<String, List<Object>> map : list) {
			// 遍历map获取表头字段，并将表头字段放进String型的数组
			Set<String> key = map.keySet();
			String titles = "";
			int count = 0;
			for (Iterator<String> it = key.iterator(); it.hasNext();) {
				if (count != 0) {
					titles += ";";
				}
				titles += (String) it.next();
				count++;
			}
			String[] titleArray = titles.split(";");

			// 第一行写入表头
			xssf_w_row = xssf_w_sheet.createRow(titleFlag);
			for (int i = 0; i < titleArray.length; i++) {
				xssf_w_cell = xssf_w_row.createCell(i);
				xssf_w_cell.setCellType(XSSFCell.CELL_TYPE_STRING);
				xssf_w_cell.setCellValue(titleArray[i]);
				xssf_w_cell.setCellStyle(head_cellStyle);
				// xssf_w_sheet.autoSizeColumn(( short ) i );
			}

			// 循环写入表数据
			for (int i = 0; i < map.get(titleArray[0]).size(); i++) {
				// 定义数据行
				xssf_w_row = xssf_w_sheet.createRow(i + titleFlag + 1);
				for (int j = 0; j < titleArray.length; j++) {
					xssf_w_cell = xssf_w_row.createCell(j);
					Object in = map.get(titleArray[j]).get(i);
					type4ExcelXLSX(in, xssf_w_cell, cellStyle_CN);
					// xssf_w_sheet.autoSizeColumn(( short ) i );
				}
			}

			// 下一个表头的写入位置，和上一个表头数据之间隔一行
			titleFlag += map.get(titleArray[0]).size() + 2;
		}

		// excel文件导出
		xssf_w_book.write(os);
		os.close();
		request.getSession().setAttribute("EXCEL_FINISH", "1");
	}

	/**
	 * 导出Excel
	 * 
	 * @param response
	 * @param request
	 * @param title
	 * @param map
	 *            key为标题，list为数据，统计信息在一行
	 * @param map
	 *            key为标题，list为数据，表数据
	 * @param type
	 *            0为xls，1为xlsx
	 * @param add
	 *            null不加结尾说明，非null时表示添加结尾说明
	 */
	public static void output(HttpServletResponse response, HttpServletRequest request, String title,
			Map<String, List<Object>> mapTitle, Map<String, List<Object>> map, int type, String add) throws Exception {

		if (type != 0 && type != 1) {
			throw new Exception("无效的excel导出类型，type=0表示xls，type=1表示xlsx");
		}
		List<Map<String, List<Object>>> listTitle = new ArrayList<Map<String, List<Object>>>();
		if (null != mapTitle)
			listTitle.add(mapTitle);
		List<Map<String, List<Object>>> list = new ArrayList<Map<String, List<Object>>>();
		list.add(map);
		if (type == 0) {
			outputXLS(response, request, title, listTitle, list, add);
		} else if (type == 1) {
			outputXLSX(response, request, title, listTitle, list, add);
		}
	}

	public static void outputXLS(HttpServletResponse response, HttpServletRequest request, String title,
			List<Map<String, List<Object>>> listTitle, List<Map<String, List<Object>>> list, String add)
					throws Exception {

		// 输出流定义
		OutputStream os = response.getOutputStream();
		byte[] fileNameByte = (title + ".xls").getBytes("GBK");
		String filename = new String(fileNameByte, "ISO8859-1");
		response.setContentType("application/x-msdownload");
		response.setCharacterEncoding("UTF-8");
		response.setHeader("Content-Disposition", "attachment;filename=" + filename);

		// 创建excel文件
		HSSFWorkbook hssf_w_book = new HSSFWorkbook();
		HSSFSheet hssf_w_sheet = hssf_w_book.createSheet(title);
		hssf_w_sheet.setDefaultColumnWidth(21); // 固定列宽度
		HSSFRow hssf_w_row = null;// 创建一行
		HSSFCell hssf_w_cell = null;// 创建每个单元格

		// 定义表头单元格样式
		HSSFCellStyle head_cellStyle = hssf_w_book.createCellStyle();
		// 定义表头字体样式
		HSSFFont head_font = hssf_w_book.createFont();
		head_font.setFontName("宋体");// 设置头部字体为宋体
		head_font.setBoldweight(Font.BOLDWEIGHT_BOLD); // 粗体
		head_font.setFontHeightInPoints((short) 10); // 字体大小
		// 表头单元格样式设置
		head_cellStyle.setFont(head_font);// 单元格样式使用字体
		head_cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		head_cellStyle.setVerticalAlignment(HSSFCellStyle.ALIGN_CENTER);
		head_cellStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		head_cellStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		head_cellStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);
		head_cellStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);

		// 定义数据单元格样式
		HSSFCellStyle cellStyle_CN = hssf_w_book.createCellStyle();// 创建数据单元格样式(数据库数据样式)
		cellStyle_CN.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		cellStyle_CN.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		cellStyle_CN.setBorderRight(XSSFCellStyle.BORDER_THIN);
		cellStyle_CN.setBorderTop(XSSFCellStyle.BORDER_THIN);

		// 在多表头导出时，定义第一个表头出现位置
		int titleFlag = 0;

		List<List<Map<String, List<Object>>>> listDatas = new ArrayList<List<Map<String, List<Object>>>>();
		if (null != listTitle)
			listDatas.add(listTitle);
		listDatas.add(list);

		for (List<Map<String, List<Object>>> listData : listDatas) {

			// 遍历写入表数据的list
			for (Map<String, List<Object>> map : listData) {
				// 遍历map获取表头字段，并将表头字段放进String型的数组
				Set<String> key = map.keySet();
				String titles = "";
				int count = 0;
				for (Iterator<String> it = key.iterator(); it.hasNext();) {
					if (count != 0) {
						titles += ";";
					}
					titles += (String) it.next();
					count++;
				}
				String[] titleArray = titles.split(";");

				// 表头写入位置
				hssf_w_row = hssf_w_sheet.createRow(titleFlag);
				for (int i = 0; i < titleArray.length; i++) {
					hssf_w_cell = hssf_w_row.createCell(i);
					hssf_w_cell.setCellType(XSSFCell.CELL_TYPE_STRING);
					hssf_w_cell.setCellValue(titleArray[i]);
					hssf_w_cell.setCellStyle(head_cellStyle);
					// hssf_w_sheet.autoSizeColumn(( short ) i );
				}

				// 循环写入表数据，获取表的总列数，然后逐行写入数据
				for (int i = 0; i < map.get(titleArray[0]).size(); i++) {
					// 定义数据行
					hssf_w_row = hssf_w_sheet.createRow(i + titleFlag + 1);
					// 按行将每一列的数据写入单元格
					for (int j = 0; j < titleArray.length; j++) {
						hssf_w_cell = hssf_w_row.createCell(j);
						Object in = map.get(titleArray[j]).get(i);
						type4ExcelXLS(in, hssf_w_cell, cellStyle_CN);
						// hssf_w_sheet.autoSizeColumn(( short ) i );
					}
				}
				// 下一个表头的写入位置，和上一个表头数据之间隔一行
				titleFlag += map.get(titleArray[0]).size() + 1;
			}
		}

		if (null != add) {
			// 定义表头单元格样式
			HSSFCellStyle head_cellStyle2 = hssf_w_book.createCellStyle();
			// 表头单元格样式设置
			head_cellStyle2.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
			head_cellStyle2.setAlignment(HSSFCellStyle.ALIGN_LEFT);
			head_cellStyle2.setWrapText(true);

			// 定义表头字体样式
			HSSFFont head_font2 = hssf_w_book.createFont();
			head_font2.setFontName("宋体");// 设置头部字体为宋体
			head_font2.setBoldweight(Font.BOLDWEIGHT_BOLD); // 粗体
			head_font2.setFontHeightInPoints((short) 9); // 字体大小
			head_font2.setColor(Font.COLOR_RED);

			HSSFFont head_font3 = hssf_w_book.createFont();
			head_font3.setFontName("宋体");// 设置头部字体为宋体
			head_font3.setFontHeightInPoints((short) 9); // 字体大小
			head_font3.setColor(Font.COLOR_RED);

			HSSFRichTextString ts = new HSSFRichTextString(add);
			ts.applyFont(0, 30, head_font2);
			ts.applyFont(30, ts.length(), head_font3);

			titleFlag += 3;
			hssf_w_row = hssf_w_sheet.createRow(titleFlag);
			hssf_w_cell = hssf_w_row.createCell(0);
			hssf_w_cell.setCellType(XSSFCell.CELL_TYPE_STRING);

			hssf_w_cell.setCellValue(ts);
			hssf_w_cell.setCellStyle(head_cellStyle2);
			hssf_w_sheet.addMergedRegion(new CellRangeAddress(titleFlag, titleFlag + 8, 0, 4));
		}
		// excel文件导出
		hssf_w_book.write(os);
		os.close();
	}

	public static void outputXLSX(HttpServletResponse response, HttpServletRequest request, String title,
			List<Map<String, List<Object>>> listTitle, List<Map<String, List<Object>>> list, String add)
					throws Exception {

		// 输出流定义
		OutputStream os = response.getOutputStream();
		byte[] fileNameByte = (title + ".xlsx").getBytes("GBK");
		String filename = new String(fileNameByte, "ISO8859-1");
		response.setContentType("application/x-msdownload");
		response.setCharacterEncoding("UTF-8");
		response.setHeader("Content-Disposition", "attachment;filename=" + filename);

		// 创建excel文件
		XSSFWorkbook xssf_w_book = new XSSFWorkbook();
		XSSFSheet xssf_w_sheet = xssf_w_book.createSheet(title);
		xssf_w_sheet.setDefaultColumnWidth(21); // 固定列宽度
		XSSFRow xssf_w_row = null;// 创建一行
		XSSFCell xssf_w_cell = null;// 创建每个单元格

		// 定义表头单元格样式
		XSSFCellStyle head_cellStyle = xssf_w_book.createCellStyle();
		// 定义表头字体样式
		XSSFFont head_font = xssf_w_book.createFont();
		head_font.setFontName("宋体");// 设置头部字体为宋体
		head_font.setBoldweight(Font.BOLDWEIGHT_BOLD); // 粗体
		head_font.setFontHeightInPoints((short) 10);
		// 表头单元格样式设置
		head_cellStyle.setFont(head_font);// 单元格使用表头字体样式
		head_cellStyle.setAlignment(HorizontalAlignment.CENTER);
		head_cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		head_cellStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		head_cellStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		head_cellStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);
		head_cellStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);

		// 定义数据单元格样式
		XSSFCellStyle cellStyle_CN = xssf_w_book.createCellStyle();// 创建数据单元格样式(数据库数据样式)
		cellStyle_CN.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		cellStyle_CN.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		cellStyle_CN.setBorderRight(XSSFCellStyle.BORDER_THIN);
		cellStyle_CN.setBorderTop(XSSFCellStyle.BORDER_THIN);

		// 在多表头导出时，定义第一个表头出现位置
		int titleFlag = 0;

		// 遍历写入表数据的list
		for (Map<String, List<Object>> map : list) {
			// 遍历map获取表头字段，并将表头字段放进String型的数组
			Set<String> key = map.keySet();
			String titles = "";
			int count = 0;
			for (Iterator<String> it = key.iterator(); it.hasNext();) {
				if (count != 0) {
					titles += ";";
				}
				titles += (String) it.next();
				count++;
			}
			String[] titleArray = titles.split(";");

			// 第一行写入表头
			xssf_w_row = xssf_w_sheet.createRow(titleFlag);
			for (int i = 0; i < titleArray.length; i++) {
				xssf_w_cell = xssf_w_row.createCell(i);
				xssf_w_cell.setCellType(XSSFCell.CELL_TYPE_STRING);
				xssf_w_cell.setCellValue(titleArray[i]);
				xssf_w_cell.setCellStyle(head_cellStyle);
				// xssf_w_sheet.autoSizeColumn(( short ) i );
			}

			// 循环写入表数据
			for (int i = 0; i < map.get(titleArray[0]).size(); i++) {
				// 定义数据行
				xssf_w_row = xssf_w_sheet.createRow(i + titleFlag + 1);
				for (int j = 0; j < titleArray.length; j++) {
					xssf_w_cell = xssf_w_row.createCell(j);
					Object in = map.get(titleArray[j]).get(i);
					type4ExcelXLSX(in, xssf_w_cell, cellStyle_CN);
					// xssf_w_sheet.autoSizeColumn(( short ) i );
				}
			}

			// 下一个表头的写入位置，和上一个表头数据之间隔一行
			titleFlag += map.get(titleArray[0]).size() + 2;
		}
		if (null != add) {
			titleFlag += 3;
			xssf_w_row = xssf_w_sheet.createRow(titleFlag);
			xssf_w_cell = xssf_w_row.createCell(0);
			xssf_w_cell.setCellType(XSSFCell.CELL_TYPE_STRING);
			xssf_w_cell.setCellValue(add);
			xssf_w_cell.setCellStyle(head_cellStyle);
			xssf_w_sheet.addMergedRegion(new CellRangeAddress(titleFlag, titleFlag + 8, 0, 4));
		}
		// excel文件导出
		xssf_w_book.write(os);
		os.close();
	}

	/**
	 * 根据类型自适应格式
	 * 
	 * @param col
	 * @param row
	 * @param in
	 * @return
	 * @throws Exception
	 */
	public static void type4ExcelXLSX(Object in, XSSFCell cell, XSSFCellStyle style) throws Exception {
		if (null == in) {
			cell.setCellType(XSSFCell.CELL_TYPE_STRING);
			XSSFRichTextString xssfString = new XSSFRichTextString("");
			cell.setCellValue(xssfString);
			cell.setCellStyle(style);
		} else {
//			in = ClobUtils.clobToString(in);
			String type = in.getClass().getName();
			if (INT.equals(type)) {
				cell.setCellType(XSSFCell.CELL_TYPE_NUMERIC);
				cell.setCellValue(Double.parseDouble(String.valueOf(in)));
				cell.setCellStyle(style);
			} else if (LONG.equals(type) && String.valueOf(in).length() <= 11) {
				cell.setCellType(XSSFCell.CELL_TYPE_NUMERIC);
				cell.setCellValue(Double.parseDouble(String.valueOf(in)));
				cell.setCellStyle(style);
			} else if (SHORT.equals(type)) {
				cell.setCellType(XSSFCell.CELL_TYPE_NUMERIC);
				cell.setCellValue(Double.parseDouble(String.valueOf(in)));
				cell.setCellStyle(style);
			} else if (DATE.equals(type)) {
				java.sql.Timestamp sqlDate = (java.sql.Timestamp) in;
				Date d = new java.util.Date(sqlDate.getTime());
				Date ds = new SimpleDateFormat("yyyy-MM-dd").parse(new SimpleDateFormat("yyyy-MM-dd").format(d));
				SimpleDateFormat df = null;
				if (d.compareTo(ds) == 0) {
					df = new SimpleDateFormat("yyyy-MM-dd");
				} else {
					df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
				}
				cell.setCellType(XSSFCell.CELL_TYPE_STRING);
				XSSFRichTextString xssfString = new XSSFRichTextString(df.format(d));
				cell.setCellValue(xssfString);
				cell.setCellStyle(style);
			} else if (in instanceof java.util.Date) {
				Date d = (Date) in;
				Date ds = new SimpleDateFormat("yyyy-MM-dd").parse(new SimpleDateFormat("yyyy-MM-dd").format(d));
				SimpleDateFormat df = null;
				if (d.compareTo(ds) == 0) {
					df = new SimpleDateFormat("yyyy-MM-dd");
				} else {
					df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
				}
				cell.setCellType(XSSFCell.CELL_TYPE_STRING);
				XSSFRichTextString xssfString = new XSSFRichTextString(df.format(d));
				cell.setCellValue(xssfString);
				cell.setCellStyle(style);
			} else if (STRING.equals(type)) {
				cell.setCellType(XSSFCell.CELL_TYPE_STRING);
				XSSFRichTextString xssfString = new XSSFRichTextString(String.valueOf(in));
				cell.setCellValue(xssfString);
				cell.setCellStyle(style);
			} else if (in instanceof BigDecimal) {
				cell.setCellType(XSSFCell.CELL_TYPE_NUMERIC);
				cell.setCellValue(Double.parseDouble(String.valueOf(in)));
				cell.setCellStyle(style);
			} else {
				try {
					double d = Double.parseDouble(String.valueOf(in));
					// if (String.valueOf(d).equals(String.valueOf(in)) &&
					// String.valueOf(in).length() <= 11){
					if (String.valueOf(in).length() <= 11) {
						cell.setCellType(XSSFCell.CELL_TYPE_NUMERIC);
						cell.setCellValue(d);
						cell.setCellStyle(style);
					} else {
						cell.setCellType(XSSFCell.CELL_TYPE_STRING);
						XSSFRichTextString xssfString = new XSSFRichTextString(String.valueOf(in));
						cell.setCellValue(xssfString);
						cell.setCellStyle(style);
					}

				} catch (Exception e) {
					cell.setCellType(XSSFCell.CELL_TYPE_STRING);
					XSSFRichTextString xssfString = new XSSFRichTextString("");
					cell.setCellValue(xssfString);
					cell.setCellStyle(style);
					LOGGER.error("excel解析:", e);
				}

			}
		}

	}

	/**
	 * 根据类型自适应格式
	 * 
	 * @param col
	 * @param row
	 * @param in
	 * @return
	 * @throws Exception
	 */
	public static void type4ExcelXLS(Object in, HSSFCell cell, HSSFCellStyle style) throws Exception {
		if (null == in) {
			cell.setCellType(HSSFCell.CELL_TYPE_STRING);
			HSSFRichTextString hssfString = new HSSFRichTextString("");
			cell.setCellValue(hssfString);
			cell.setCellStyle(style);
		} else {
//			in = clobToString(in);
			String type = in.getClass().getName();
			if (INT.equals(type)) {
				cell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);
				cell.setCellValue(Double.parseDouble(String.valueOf(in)));
				cell.setCellStyle(style);
			} else if (LONG.equals(type) && String.valueOf(in).length() <= 11) {
				cell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);
				cell.setCellValue(Double.parseDouble(String.valueOf(in)));
				cell.setCellStyle(style);
			} else if (SHORT.equals(type)) {
				cell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);
				cell.setCellValue(Double.parseDouble(String.valueOf(in)));
				cell.setCellStyle(style);
			} else if (DATE.equals(type)) {
				java.sql.Timestamp sqlDate = (java.sql.Timestamp) in;
				Date d = new java.util.Date(sqlDate.getTime());
				Date ds = new SimpleDateFormat("yyyy-MM-dd").parse(new SimpleDateFormat("yyyy-MM-dd").format(d));
				SimpleDateFormat df = null;
				if (d.compareTo(ds) == 0) {
					df = new SimpleDateFormat("yyyy-MM-dd");
				} else {
					df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
				}
				cell.setCellType(HSSFCell.CELL_TYPE_STRING);
				HSSFRichTextString hssfString = new HSSFRichTextString(df.format(d));
				cell.setCellValue(hssfString);
				cell.setCellStyle(style);
			} else if (in instanceof java.util.Date) {
				Date d = (Date) in;
				Date ds = new SimpleDateFormat("yyyy-MM-dd").parse(new SimpleDateFormat("yyyy-MM-dd").format(d));
				SimpleDateFormat df = null;
				if (d.compareTo(ds) == 0) {
					df = new SimpleDateFormat("yyyy-MM-dd");
				} else {
					df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
				}
				cell.setCellType(HSSFCell.CELL_TYPE_STRING);
				HSSFRichTextString hssfString = new HSSFRichTextString(df.format(d));
				cell.setCellValue(hssfString);
				cell.setCellStyle(style);
			} else if (STRING.equals(type)) {
				cell.setCellType(HSSFCell.CELL_TYPE_STRING);
				HSSFRichTextString hssfString = new HSSFRichTextString(String.valueOf(in));
				cell.setCellValue(hssfString);
				cell.setCellStyle(style);
			} else if (in instanceof BigDecimal) {
				cell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);
				cell.setCellValue(Double.parseDouble(String.valueOf(in)));
				cell.setCellStyle(style);
			} else {
				try {
					double d = Double.parseDouble(String.valueOf(in));
					// if (String.valueOf(d).equals(String.valueOf(in)) &&
					// String.valueOf(in).length() <= 11){
					if (String.valueOf(in).length() <= 11) {
						cell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);
						cell.setCellValue(d);
						cell.setCellStyle(style);
					} else {
						cell.setCellType(HSSFCell.CELL_TYPE_STRING);
						HSSFRichTextString hssfString = new HSSFRichTextString(String.valueOf(in));
						cell.setCellValue(hssfString);
						cell.setCellStyle(style);
					}

				} catch (Exception e) {
					cell.setCellType(HSSFCell.CELL_TYPE_STRING);
					HSSFRichTextString hssfString = new HSSFRichTextString("");
					cell.setCellValue(hssfString);
					cell.setCellStyle(style);
					LOGGER.error("excel解析:", e);
				}

			}
		}

	}
	
	 /** 
     * 数据库Clob对象转换为String 
     */  
    @SuppressWarnings("unused")  
    public static String clobToString(Clob clob)  
    {  
        if(clob == null) {  
            return null;  
        }  
        try  
        {  
            Reader inStreamDoc = clob.getCharacterStream();  
              
            char[] tempDoc = new char[(int) clob.length()];  
            inStreamDoc.read(tempDoc);  
            inStreamDoc.close();  
            return new String(tempDoc);  
        }  
        catch (IOException e)  
        {  
            e.printStackTrace();  
        }  
        catch (SQLException es)  
        {  
            es.printStackTrace();  
        }  
          
        return null;  
    }  
}