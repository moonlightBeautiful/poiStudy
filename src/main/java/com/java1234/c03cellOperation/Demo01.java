package com.java1234.c03cellOperation;

import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.FileOutputStream;

/**
 * 单元格位置格式，位置
 */
public class Demo01 {

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
        // 创建一个单元格 第3行第2列，关于位置的操作
        Cell cell32 = row3.createCell(1);
        cell32.setCellValue(new HSSFRichTextString("Align It"));
        // 垂直上边 水平左边
        /*cellStyle.setVerticalAlignment(VerticalAlignment.TOP);
        cellStyle.setAlignment(HorizontalAlignment.LEFT);*/
        // 垂直居中 水平居中
        /*cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);*/
        // 垂直下边 水平右边
        /*cellStyle.setVerticalAlignment(VerticalAlignment.BOTTOM);
        cellStyle.setAlignment(HorizontalAlignment.RIGHT);*/
        // 垂直居中 水平两边
        cellStyle.setVerticalAlignment(VerticalAlignment.DISTRIBUTED);
        cellStyle.setAlignment(HorizontalAlignment.DISTRIBUTED);
        cell32.setCellStyle(cellStyle);

        FileOutputStream fileOut = new FileOutputStream("c:\\测试工作簿.xls");
        wb.write(fileOut);
        fileOut.close();
    }

}
