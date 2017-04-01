package com.bucuoa.common.excel;

import java.io.IOException;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.Map.Entry;

import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellRangeAddress;

public class ExcelExportUtils {
	
//	Map<String,String> meta = new LinkedHashMap<String,String>();
//	meta.put("productName", "商品名称");
//	meta.put("sku_id", "sku");

	
	public static void download(HttpServletResponse response, List<Map<String, Object>> searchDataList,
			Map<String, String> meta,String filename,String sheetname) {
		HSSFWorkbook wb = new HSSFWorkbook();  
		
		//创建HSSFSheet对象  
		HSSFSheet sheet = wb.createSheet(sheetname);  
		//创建HSSFRow对象  
		HSSFRow rowtitle = sheet.createRow(0);  
		//创建HSSFCell对象  
		Set<Entry<String, String>> entrySet1 = meta.entrySet();
		int k = 0;
		for(Entry<String, String> ent  : entrySet1)
		{
			HSSFCell cell = rowtitle.createCell(k);  
			//设置单元格的值  
				try {
					if(ent.getKey()!=null)
					{
						String key = ent.getKey().toString();
						cell.setCellValue(meta.get(key));
					}
				} catch (Exception e) {
					e.printStackTrace();
				}  
			k ++;
		}
		
		for( int i = 0 ; i < searchDataList.size() ; i ++)
		{
			
			HSSFRow row = sheet.createRow(i+1);  
			//创建HSSFCell对象  
			Map<String, Object> map = searchDataList.get(i);
			Set<Entry<String, String>> entrySet2 = meta.entrySet();
			int j = 0;
			for(Entry<String, String> ent  : entrySet2)
			{
				HSSFCell cell = row.createCell(j);  
				//设置单元格的值  
					try {
						Object object = map.get(ent.getKey());
						if(object!=null)
						{
							cell.setCellValue(object.toString());
						}
					} catch (Exception e) {
						e.printStackTrace();
					}  
				j ++;
			}
		}

		String codedFileName="";
		try {
			codedFileName = java.net.URLEncoder.encode(filename, "UTF-8");
		} catch (UnsupportedEncodingException e) {
			e.printStackTrace();
		}
	
			response.setContentType("application/vnd.ms-excel");    
	        response.setHeader("Content-disposition", "attachment;filename="+codedFileName);    
	        try {
				OutputStream ouputStream = response.getOutputStream();    
				wb.write(ouputStream);    
				ouputStream.flush();    
				ouputStream.close();
			} catch (IOException e1) {
				e1.printStackTrace();
			}
	}

}
