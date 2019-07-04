package com.java1234.c01hello;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

/**
 * @author gaoxu
 * @date 2019-07-04 17:31
 * @description ... 类
 */
public class Hello {

    public static void main(String[] args) {

        try {
            // poi创建一个工作簿
            Workbook wb = new HSSFWorkbook();
            // 创建第一个Sheet页
            Sheet sheet1 = wb.createSheet("第一个Sheet页");
            // 创建第二个Sheet页
            Sheet sheet2 = wb.createSheet("第二个Sheet页");
            // 第一个Sheet页中创建第一行
            Row sheet1Row1 = sheet1.createRow(0);
            // 第一个Sheet页第一行中创建从左到右创建4个单元格，设置值
            Cell sheet1Row1cell1 = sheet1Row1.createCell(0);
            sheet1Row1cell1.setCellValue(1);
            Cell sheet1Row1cell2 = sheet1Row1.createCell(1);
            sheet1Row1cell2.setCellValue(1.2);
            Cell sheet1Row1cell3 = sheet1Row1.createCell(2);
            sheet1Row1cell3.setCellValue("这是一个字符串类型");
            Cell sheet1Row1cell4 = sheet1Row1.createCell(3);
            sheet1Row1cell4.setCellValue(false);
            // 在指定位置写入工作簿
            OutputStream fileOut = new FileOutputStream("c:\\用Poi搞出来的工作簿.xls");
            wb.write(fileOut);
            fileOut.close();
        } catch (IOException e) {
            e.printStackTrace();
        }


    }
}
