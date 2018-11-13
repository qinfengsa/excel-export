package com.qinfengsama.utils;

import java.awt.Color;
import java.io.File;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.List;
import java.util.Map;
import java.util.UUID;

import javax.servlet.http.HttpServletRequest;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtils {
	
	/**
	 * 创建Excel文件，并返回excel文件路径
	 * @param list 
	 * @param tableHeader
	 * @param tableProperties
	 * @param userId //用户id 保证文件名不重复
	 * @param request
	 * @return
	 * @throws Exception
	 */
	public static String createExcelSheetUrl(List<Map<String,Object>> list,String[] tableHeader,
			String[] tableProperties,String userId,HttpServletRequest request) throws Exception {
		String url = ""; 
		SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		XSSFWorkbook workbook = new XSSFWorkbook();
        //XSSFCellStyle style = wb.createCellStyle();   
        XSSFSheet sheet = workbook.createSheet("sheet1");
   
 
        XSSFRow headerRow = sheet.createRow(0); 
        XSSFCellStyle headstyle = workbook.createCellStyle();
        
        XSSFFont headfont = workbook.createFont();
        //在对应的workbook中新建字体
        headfont.setFontName("微软雅黑");
        headfont.setBold(true); 
    
        headstyle.setFont(headfont);
        XSSFCell headerCell = null; //在循环外创建对象，保证内存不会溢出
        for (int i = 0; i < tableHeader.length; i++) {//标题行
            headerCell = headerRow.createCell(i);
            headerCell.setCellStyle(headstyle);
            // 设置cell的值
            headerCell.setCellValue(tableHeader[i]);
            headerCell.setCellStyle(headstyle);
        }
        //在循环外创建对象，保证内存不会溢出
        XSSFRow row = null;
        XSSFCell cell = null; 
        Map<String,Object> map = null;
        int rowIndex = 1;
        for (int i = 0 ; i < list.size() ; i++) {
            map = (Map<String,Object>) list.get(i);
            row = sheet.createRow(rowIndex);//创建一行数据
            for (int j = 0; j < tableProperties.length; j++) {//循环放入每格的数据
                // 创建第i个单元格
                cell = row.createCell(j);
                if(map.containsKey(tableProperties[j]) && 
                        map.get(tableProperties[j]) != null && 
                        map.get(tableProperties[j]) != "" && 
                        (map.get(tableProperties[j]).getClass().equals(java.util.Date.class) 
                                || map.get(tableProperties[j]).getClass().equals(java.sql.Date.class)
                                || map.get(tableProperties[j]).getClass().equals(java.sql.Timestamp.class)) ){
                    map.put(tableProperties[j], dateFormat.format(map.get(tableProperties[j])));
                }
                cell.setCellValue(map.get(tableProperties[j]) == null ?"":map.get(tableProperties[j]).toString());
                sheet.setColumnWidth(j, (1000 * 5));
            }
            map.clear();
            rowIndex++;
        }
        
        //采用UUID的方式   ，可以产生不同的字符变量的 保证唯一命名
        UUID randomUUID = UUID.randomUUID();//j
 
   	    String xlsName = userId + "_" + randomUUID.toString() + ".xlsx";
   	    String xlsPath = "/download/excel/" + xlsName; 
   	    
   	    String realPath = request.getSession().getServletContext().getRealPath("/") + xlsPath; 
		
	   	File staticFile = new File(realPath);
   	    File staticDirectory = staticFile.getParentFile();
   	    if (!staticDirectory.exists()) {//判断 父文件夹是否存在    
   	    	staticDirectory.mkdirs();
   	    } 
   	    FileOutputStream os = new FileOutputStream(realPath);    
	 
		workbook.write(os);
		os.close();
		
		url = xlsPath;

		 
		return url ; 
	}
	
}
