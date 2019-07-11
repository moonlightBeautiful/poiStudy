package com.java1234.c03cellOperation;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.FileOutputStream;

/**
 * 单元格位置格式，但是以前的方法都废弃了
 */
public class Demo01 {

    public static void main(String[] args) throws Exception {
        // 创建一个新的工作簿
        Workbook wb = new HSSFWorkbook();
        // 创建Sheet页，第一页
        Sheet sheet = wb.createSheet("第一个Sheet页");
        // 创建一个行，第三行
        Row row = sheet.createRow(2);
        // 设置行高30
        row.setHeightInPoints(30);

        createCell(wb, row, (short) 0, HSSFCellStyle.class., HSSFCellStyle.VERTICAL_BOTTOM);
        createCell(wb, row, (short) 1, HSSFCellStyle.ALIGN_FILL, HSSFCellStyle.VERTICAL_CENTER);
        createCell(wb, row, (short) 2, HSSFCellStyle.ALIGN_LEFT, HSSFCellStyle.VERTICAL_TOP);
        createCell(wb, row, (short) 3, HSSFCellStyle.ALIGN_RIGHT, HSSFCellStyle.VERTICAL_TOP);

        FileOutputStream fileOut = new FileOutputStream("c:\\������.xls");
        wb.write(fileOut);
        fileOut.close();
    }

    /**
     * ����һ����Ԫ��Ϊ���趨ָ���Ķ��䷽ʽ
     *
     * @param wb     ������
     * @param row    ��
     * @param column ��
     * @param halign ˮƽ������䷽ʽ
     * @param valign ��ֱ������䷽ʽ
     */
    private static void createCell(Workbook wb, Row row, short column, short halign, short valign) {
        // 创建单元格
        Cell cell = row.createCell(column);
        // 设置值
        cell.setCellValue(new HSSFRichTextString("Align It"));
        CellStyle cellStyle = wb.createCellStyle();
        //
        HorizontalAlignment
        cellStyle.setAlignment(halign);
        cellStyle.setVerticalAlignment(valign);
        cell.setCellStyle(cellStyle);
    }


}
