package com.example.excel;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class ExcelApplication {
    
	public static void main(String[] args) {
            ExcelDataHandler dataHandler = new ExcelDataHandler();
            
            XSSFWorkbook workbook = dataHandler.createProductRankingWorkbook();
            
            try {
                FileOutputStream fileOut = new FileOutputStream("Product Ranking.xlsx");
                workbook.write(fileOut);
                fileOut.close();
            } catch (FileNotFoundException ex) {
                System.out.println(ex.getCause());
            } catch (IOException ex) {
                System.out.println(ex.getCause());
            }
            
            
	}

}
