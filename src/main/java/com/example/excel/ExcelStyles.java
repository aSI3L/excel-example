package com.example.excel;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelStyles {
    
        public XSSFCellStyle headerStyles(XSSFWorkbook workbook) {
            XSSFCellStyle headerStyle = workbook.createCellStyle();
            headerStyle.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            
            XSSFFont headerFont = workbook.createFont();
            headerFont.setFontName("Times New Roman");
            headerFont.setColor(IndexedColors.BLACK.getIndex());
            headerFont.setBold(true);
            headerStyle.setFont(headerFont);
            
            headerStyle.setAlignment(HorizontalAlignment.CENTER);
            headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            
            headerStyle.setBorderTop(BorderStyle.THIN);
            headerStyle.setBorderBottom(BorderStyle.THIN);
            headerStyle.setBorderLeft(BorderStyle.THIN);
            headerStyle.setBorderRight(BorderStyle.THIN);
            
            return headerStyle;
        }
        
        public XSSFCellStyle dataStyles(XSSFWorkbook workbook) {
            XSSFCellStyle dataStyle = workbook.createCellStyle();
            dataStyle.setFillForegroundColor(IndexedColors.LEMON_CHIFFON.getIndex());
            dataStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            
            XSSFFont dataFont = workbook.createFont();
            dataFont.setFontName("Times New Roman");
            dataFont.setColor(IndexedColors.BLACK.getIndex());
            dataStyle.setFont(dataFont);
            
            dataStyle.setAlignment(HorizontalAlignment.CENTER);
            dataStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            
            dataStyle.setBorderTop(BorderStyle.THIN);
            dataStyle.setBorderBottom(BorderStyle.THIN);
            dataStyle.setBorderLeft(BorderStyle.THIN);
            dataStyle.setBorderRight(BorderStyle.THIN);
            
            return dataStyle;
        }
        
        public XSSFCellStyle statusStyles(XSSFWorkbook workbook, XSSFCellStyle stylesToClone, Boolean active) {
            XSSFFont activeFont = workbook.createFont();
            activeFont.setFontName("Times New Roman");
            activeFont.setColor(IndexedColors.BLACK.getIndex());
            
            XSSFFont unactiveFont = workbook.createFont();
            unactiveFont.setFontName("Times New Roman");
            unactiveFont.setColor(IndexedColors.WHITE.getIndex());
            
            XSSFCellStyle activeCellStyle = workbook.createCellStyle();
            activeCellStyle.cloneStyleFrom(stylesToClone);
            
            if(active) {
                activeCellStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
                activeCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                activeCellStyle.setFont(activeFont);
            } else {
                activeCellStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
                activeCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                activeCellStyle.setFont(unactiveFont);
            }
            
            return activeCellStyle;
            
        }
}
