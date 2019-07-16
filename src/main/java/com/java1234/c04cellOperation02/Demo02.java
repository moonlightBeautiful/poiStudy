package com.java1234.c04cellOperation02;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;

public class Demo02 {

    public static void main(String[] args) throws Exception {
        InputStream inp = new FileInputStream("c:\\工作簿.xls");
        POIFSFileSystem fs = new POIFSFileSystem(inp);
        //读取Excel文件
        Workbook wb = new HSSFWorkbook(fs);
        // 获取第一个Sheet页
        Sheet sheet = wb.getSheetAt(0);
        // 获取第一行
        Row row = sheet.getRow(0);
        // 获取第一个单元格

        Cell cell = row.getCell(0, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
        //单元格没有值时，getCell方法获取不到单元格，可能需要指定MissingCellPolicy,但是指定了也没用
        if (cell == null) {
            cell = row.createCell(3);
        }
        cell.setCellType(Cell.CELL_TYPE_STRING);
        cell.setCellValue("测试单元格");

        FileOutputStream fileOut = new FileOutputStream("c:\\工作簿.xls");
        wb.write(fileOut);
        fileOut.close();
    }
}
