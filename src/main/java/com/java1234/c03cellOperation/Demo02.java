/*
package com.java1234.c03cellOperation;

import java.io.FileOutputStream;
import java.util.Calendar;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

*/
/**
 * 设置单元格边框，以前的方法也废弃了
 *//*

public class Demo02 {

	public static void main(String[] args) throws Exception{
		Workbook wb=new HSSFWorkbook();
		Sheet sheet=wb.createSheet("��һ��Sheetҳ");
		Row row=sheet.createRow(1);
		
		Cell cell=row.createCell(1);
		cell.setCellValue(4);
		
		CellStyle cellStyle=wb.createCellStyle(); 
		cellStyle.setBorderBottom(CellStyle.BORDER_THIN); // �ײ��߿�
		cellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex()); // �ײ��߿���ɫ
		
		cellStyle.setBorderLeft(CellStyle.BORDER_THIN);  // ��߱߿�
		cellStyle.setLeftBorderColor(IndexedColors.GREEN.getIndex()); // ��߱߿���ɫ
		
		cellStyle.setBorderRight(CellStyle.BORDER_THIN); // �ұ߱߿�
		cellStyle.setRightBorderColor(IndexedColors.BLUE.getIndex());  // �ұ߱߿���ɫ
		
		cellStyle.setBorderTop(CellStyle.BORDER_MEDIUM_DASHED); // �ϱ߱߿�
		cellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());  // �ϱ߱߿���ɫ
		
		cell.setCellStyle(cellStyle);
		FileOutputStream fileOut=new FileOutputStream("c:\\������.xls");
		wb.write(fileOut);
		fileOut.close();
	}
}
*/
