package com.java1234.c02cellContentOperation;

import org.apache.poi.hssf.extractor.ExcelExtractor;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.FileInputStream;
import java.io.InputStream;

/**
 * @author gaoxu
 * @date 2019-07-09 14:26
 * @description ... 类
 */
public class Demo03extractSheetTextContent {
    public static void main(String[] args) throws Exception {
        InputStream is = new FileInputStream("c:\\工作簿.xls");
        POIFSFileSystem fs = new POIFSFileSystem(is);
        HSSFWorkbook wb = new HSSFWorkbook(fs);
        ExcelExtractor excelExtractor = new ExcelExtractor(wb);
        // 我们不需要Sheet页的名字
        /*excelExtractor.setIncludeSheetNames(false);*/
        System.out.println(excelExtractor.getText());
    }
}
