package com.java1234.c02cellOperation;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Calendar;
import java.util.Date;

/**
 * @author gaoxu
 * @date 2019-07-09 13:57
 * @description ... 类
 */
public class Demo01cellData {
    public static void main(String[] args) throws IOException {
        // 定义一个新的工作簿
        Workbook wb = new HSSFWorkbook();
        //单元格样式类
        CellStyle cellStyle = wb.createCellStyle();
        CreationHelper createHelper = wb.getCreationHelper();
        // 指定时间格式化样式
        cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("yyy-mm-dd hh:mm:ss"));

        // 创建第一个Sheet页
        Sheet sheet = wb.createSheet("第一个Sheet页");
        // 创建第一行
        Row rowL1 = sheet.createRow(0);


        //第1行1列：时间
        rowL1.createCell(0).setCellValue(new Date());

        //第1行2列：时间
        Cell cell12 = rowL1.createCell(1);
        cell12.setCellValue(new Date());
        cell12.setCellStyle(cellStyle);

        //第1行3列：时间
        Cell cell13 = rowL1.createCell(2);
        cell13.setCellValue(Calendar.getInstance());
        cell13.setCellStyle(cellStyle);

        //第1行4列：数值
        rowL1.createCell(3).setCellValue(1);

        //第1行5列：数值
        rowL1.createCell(4).setCellValue(1.1);

        //第1行6列：字符串
        rowL1.createCell(5).setCellValue("一个字符串");

        //第1行7列：boolean
        rowL1.createCell(6).setCellValue(true);

        // 工作簿写入到硬盘
        FileOutputStream fileOut = new FileOutputStream("c:\\工作簿.xls");
        wb.write(fileOut);
        fileOut.close();
    }
}
