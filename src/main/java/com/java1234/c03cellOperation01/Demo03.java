package com.java1234.c03cellOperation01;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.FileOutputStream;

/**
 * 设置单元格背景色和前景色
 */
public class Demo03 {

    public static void main(String[] args) throws Exception {
        // 创建一个新的工作簿
        Workbook wb = new HSSFWorkbook();
        // 创建Sheet页，第1页
        Sheet sheet = wb.createSheet("第一个Sheet页");
        // 创建一个行，第3行
        Row row3 = sheet.createRow(2);
        // 设置行高30
        row3.setHeightInPoints(30);

        // 创建一个单元格 第3行第2列，关于颜色的操作
        Cell cell32 = row3.createCell(1);
        cell32.setCellValue("XXX");
        // 创建一个单元格样式
        CellStyle cellStyle1 = wb.createCellStyle();
        //背景
        cellStyle1.setFillBackgroundColor(IndexedColors.GREEN.getIndex());
        cellStyle1.setFillPattern(FillPatternType.BIG_SPOTS);
        cell32.setCellStyle(cellStyle1);

        // 创建一个单元格 第3行第4列，关于颜色的操作
        Cell cell34 = row3.createCell(3);
        cell34.setCellValue("YYY");
        CellStyle cellStyle2 = wb.createCellStyle();
        //前景
        cellStyle2.setFillForegroundColor(IndexedColors.RED.getIndex());
        cellStyle2.setFillPattern(FillPatternType.BIG_SPOTS);
        cell34.setCellStyle(cellStyle2);

        FileOutputStream fileOut = new FileOutputStream("c:\\测试工作簿.xls");
        wb.write(fileOut);
        fileOut.close();
    }
}
