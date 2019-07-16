package com.java1234.c03cellOperation01;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.FileOutputStream;

/**
 * 合并单元格
 */
public class Demo04 {

    public static void main(String[] args) throws Exception {
        // 创建一个新的工作簿
        Workbook wb = new HSSFWorkbook();
        // 创建Sheet页，第1页
        Sheet sheet = wb.createSheet("第一个Sheet页");
        // 创建一个行，第3行
        Row row3 = sheet.createRow(2);
        // 设置行高30
        row3.setHeightInPoints(30);
        // 创建单元格 3行2列
        Cell cell32 = row3.createCell(1);
        cell32.setCellValue("3行2列");
        sheet.createRow(1).createCell(1).setCellValue("2行2列");
        // 合并单元格 起始行2、结束行3、起始列2、结束列4
        sheet.addMergedRegion(new CellRangeAddress(1, 2, 1, 3));

        FileOutputStream fileOut = new FileOutputStream("c:\\测试工作簿.xls");
        wb.write(fileOut);
        fileOut.close();
    }
}
