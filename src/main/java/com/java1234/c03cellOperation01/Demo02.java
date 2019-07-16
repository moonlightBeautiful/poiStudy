package com.java1234.c03cellOperation01;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.FileOutputStream;

/**
 * 设置单元格，边框
 */
public class Demo02 {

    public static void main(String[] args) throws Exception {
        // 创建一个新的工作簿
        Workbook wb = new HSSFWorkbook();
        // 创建一个单元格样式
        CellStyle cellStyle = wb.createCellStyle();
        // 创建Sheet页，第1页
        Sheet sheet = wb.createSheet("第一个Sheet页");
        // 创建一个行，第3行
        Row row3 = sheet.createRow(2);
        // 设置行高30
        row3.setHeightInPoints(30);
        // 创建一个单元格 第3行第2列，关于颜色的操作
        Cell cell32 = row3.createCell(1);

        cell32.setCellValue(4);
        //上下左右边框
        cellStyle.setBorderTop(BorderStyle.THICK);
        cellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());

        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());

        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setLeftBorderColor(IndexedColors.GREEN.getIndex());

        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setRightBorderColor(IndexedColors.BLUE.getIndex());

        cell32.setCellStyle(cellStyle);
        FileOutputStream fileOut = new FileOutputStream("c:\\测试工作簿.xls");
        wb.write(fileOut);
        fileOut.close();
    }
}
