package com.java1234.c02cellOperation;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.FileInputStream;
import java.io.InputStream;

/**
 * @author gaoxu
 * @date 2019-07-09 14:20
 * @description ... 类
 */
public class Demo02traversalCell {
    public static void main(String[] args) throws Exception {
        //读取本地xls文件为工作簿
        InputStream is = new FileInputStream("c:\\工作簿.xls");
        POIFSFileSystem fs = new POIFSFileSystem(is);
        HSSFWorkbook wb = new HSSFWorkbook(fs);
        //读取工作簿中的第一个sheet
        HSSFSheet hssfSheet = wb.getSheetAt(0);
        if (hssfSheet == null) {
            return;
        }
        // 遍历行Row
        for (int rowNum = 0; rowNum <= hssfSheet.getLastRowNum(); rowNum++) {
            HSSFRow hssfRow = hssfSheet.getRow(rowNum);
            if (hssfRow == null) {
                continue;
            }
            // 遍历列Cell
            for (int cellNum = 0; cellNum <= hssfRow.getLastCellNum(); cellNum++) {
                HSSFCell hssfCell = hssfRow.getCell(cellNum);
                if (hssfCell == null) {
                    continue;
                }
                if (hssfCell.getCellType() == HSSFCell.CELL_TYPE_BOOLEAN) {
                    System.out.print(" " + String.valueOf(hssfCell.getBooleanCellValue()));
                } else if (hssfCell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC) {
                    System.out.print(" " + String.valueOf(hssfCell.getNumericCellValue()));
                } else {
                    System.out.print(" " + String.valueOf(hssfCell.getStringCellValue()));
                }
            }
            System.out.println();
        }
    }
}
