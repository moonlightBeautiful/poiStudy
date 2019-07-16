package com.java1234.c04cellOperation02;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.FileOutputStream;

public class Demo03 {

    public static void main(String[] args) throws Exception {
        // 定义一个新的工作簿
        Workbook wb = new HSSFWorkbook();
        // 创建第一个Sheet页
        Sheet sheet = wb.createSheet("第一个Sheet页");
        // 创建一个行
        Row row = sheet.createRow(2);
        Cell cell = row.createCell(2);
        cell.setCellValue("我要换行 \n 成功了吗？");

        CellStyle cs = wb.createCellStyle();
        // 设置可以换行
        cs.setWrapText(true);
        cell.setCellStyle(cs);

        // 调整下行的高度
        row.setHeightInPoints(2 * sheet.getDefaultRowHeightInPoints());
        // 调整单元格宽度
        sheet.autoSizeColumn(2);

        FileOutputStream fileOut = new FileOutputStream("c:\\工作簿.xls");
        wb.write(fileOut);
        fileOut.close();
    }
}
