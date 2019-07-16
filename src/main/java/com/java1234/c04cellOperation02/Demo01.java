package com.java1234.c04cellOperation02;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.FileOutputStream;

public class Demo01 {

    public static void main(String[] args) throws Exception {
        // 定义一个新的工作簿
        Workbook wb = new HSSFWorkbook();
        // 创建Sheet页
        Sheet sheet = wb.createSheet("第一个Sheet页");
        // 创建行：第2行
        Row row = sheet.createRow(1);
        // 创建单元格：第2行2列
        Cell cell = row.createCell((short) 1);
        // 创建一个字体处理类
        Font font = wb.createFont();
        font.setFontHeightInPoints((short) 24); //设置字号，字的大小超过单元格大小，单元格会随之变大
        font.setFontName("Courier New");    //设置字体名称
        font.setItalic(true);       //设置字体倾斜
        font.setStrikeout(true);    //设置删除线
        /*
        font.setColor(HSSFColor.RED.index);//设置字体颜色
        font.setUnderline(FontFormatting.U_SINGLE);//设置下划线
        font.setTypeOffset(FontFormatting.SS_SUPER);//设置上标下标
        */
        CellStyle style = wb.createCellStyle();
        style.setFont(font);
        cell.setCellValue("This is test of fonts");
        cell.setCellStyle(style);

        FileOutputStream fileOut = new FileOutputStream("c:\\工作簿.xls");
        wb.write(fileOut);
        fileOut.close();
    }
}
