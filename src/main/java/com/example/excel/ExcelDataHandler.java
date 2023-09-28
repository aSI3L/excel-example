package com.example.excel;

import java.util.ArrayList;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelDataHandler {
    
        public XSSFWorkbook createProductRankingWorkbook() {
            ArrayList<Producto> products = new ArrayList<>();
            products.add(new Producto("Hamburguesa Simple", "Hamburguesas", 1500.0, true, 25));
            products.add(new Producto("Hamburguesa Completa", "Hamburguesas", 1900.0, true, 22));
            products.add(new Producto("Pancho Doble", "Panchos", 1200.0, false, 15));
            products.add(new Producto("Papas", "Fritos", 900.0, true, 12));
            
            ExcelStyles styles = new ExcelStyles();
            
            // CREATE BOOK
            XSSFWorkbook workbook = new XSSFWorkbook();
            
            // CREATE SHEETS
            XSSFSheet foodSheet = workbook.createSheet("Food Ranking");
            XSSFSheet drinkSheet = workbook.createSheet("Drinks Ranking");
            
            // CREATE HEADER ROWS
            XSSFRow foodHeaderRow = foodSheet.createRow(1);
            XSSFRow drinkHeaderRow = drinkSheet.createRow(1);
            
            // HEADER STYLES
            XSSFCellStyle headerStyle = styles.headerStyles(workbook);
            
            // DATA STYLES
            XSSFCellStyle dataStyle = styles.dataStyles(workbook);
            
            // CREATE HEADER
            createHeader(foodHeaderRow, Headers.productRankingHeaders, headerStyle);
            createHeader(drinkHeaderRow, Headers.productRankingHeaders, headerStyle);
            
            // CREATE DATA
            createProductsData(products, Headers.productRankingHeaders, workbook, foodSheet, dataStyle);
            createProductsData(products, Headers.productRankingHeaders, workbook, drinkSheet, dataStyle);
            
            // AUTO SIZE COLUMNS
            autoSizeColumns(Headers.productRankingHeaders, foodSheet);
            autoSizeColumns(Headers.productRankingHeaders, drinkSheet);
            
            return workbook;
        }
        
        public XSSFWorkbook createClientRankingWorkbook() {
            ArrayList<User> clients = new ArrayList<>();
            clients.add(new User("Pepe", "Fulano", 18, 40200.0));
            clients.add(new User("Fulano", "Mengano", 14, 27600.0));
            clients.add(new User("Pepe", "Fulano", 9, 13100.0));
            clients.add(new User("Eduardo", "Martinez", 2, 5600.0));
            
            ExcelStyles styles = new ExcelStyles();
            
            XSSFWorkbook workbook = new XSSFWorkbook();
            
            XSSFSheet clientSheet = workbook.createSheet("Client Ranking");

            XSSFRow clientHeaderRow = clientSheet.createRow(1);
            
            XSSFCellStyle headerStyle = styles.headerStyles(workbook);
            
            XSSFCellStyle dataStyle = styles.dataStyles(workbook);
            
            createHeader(clientHeaderRow, Headers.clientRankingHeaders, headerStyle);
            
            createClientsData(clients, Headers.clientRankingHeaders, clientSheet, dataStyle);
            
            autoSizeColumns(Headers.clientRankingHeaders, clientSheet);
            
            return workbook;
        }
        
        public XSSFWorkbook createMovementWorkbook() {
            ArrayList<Movements> movements = new ArrayList<>();
            movements.add(new Movements(2000, 500));
            movements.add(new Movements(800, 500));
            movements.add(new Movements(700, 200));
            movements.add(new Movements(1200, 400));
            
            ExcelStyles styles = new ExcelStyles();
            
            XSSFWorkbook workbook = new XSSFWorkbook();
            
            XSSFSheet movementSheet = workbook.createSheet("Movements");
            
            XSSFRow movementHeaderRow = movementSheet.createRow(1);
            
            XSSFCellStyle headerStyle = styles.headerStyles(workbook);
            
            XSSFCellStyle dataStyle = styles.dataStyles(workbook);
            
            createHeader(movementHeaderRow, Headers.movementsHeaders, headerStyle);
            
            createMovementsData(movements, Headers.movementsHeaders, workbook, movementSheet, dataStyle);
            
            autoSizeColumns(Headers.movementsHeaders, movementSheet);
            
            return workbook;
        }
    
        public void createHeader(XSSFRow headerRow, String [] headers, XSSFCellStyle styles) {
            for(int i = 0; i < headers.length; i++) {
                headerRow.createCell(i+1).setCellValue(headers[i]);
                XSSFCell foodHeaderCell = headerRow.getCell(i+1);
                if(headers[i].length() != 0) foodHeaderCell.setCellStyle(styles);
            }
        }
        
        public void createProductsData(ArrayList<Producto> products, String [] headers, XSSFWorkbook workbook, XSSFSheet sheet, XSSFCellStyle cellStyles) {
            ExcelStyles styles = new ExcelStyles();
            
            for(int i = 1; i <= products.size(); i++) {
                XSSFRow row = sheet.createRow(i+1);
                
                row.createCell(1).setCellValue(products.get(i-1).getName());
                row.createCell(2).setCellValue(products.get(i-1).getCategory());
                row.createCell(3).setCellValue("$" + products.get(i-1).getPrice());
                row.createCell(5).setCellValue(products.get(i-1).getQuantity_sold());           
                
                XSSFCell activeCell = row.createCell(4);
                XSSFCellStyle activeCellStyle = styles.statusStyles(workbook, cellStyles, products.get(i-1).getActive());
                
                activeCell.setCellValue(products.get(i-1).getActive() ? "Active" : "Not Active");
                
                activeCell.setCellStyle(activeCellStyle);
                
                for(int j = 1; j <= headers.length; j++) if(j != 4) row.getCell(j).setCellStyle(cellStyles);
            }
        }
        
        public void createClientsData(ArrayList<User> clients, String [] headers, XSSFSheet sheet, XSSFCellStyle cellStyles){
            for(int i = 1; i <= clients.size(); i++) {
                XSSFRow row = sheet.createRow(i+1);
                
                row.createCell(1).setCellValue(clients.get(i-1).getFirstName() + " " + clients.get(i-1).getLastName());
                row.createCell(2).setCellValue(clients.get(i-1).getOrder_quantity());
                row.createCell(3).setCellValue("$" + clients.get(i-1).getTotal_spent());
                
                for(int j = 1; j <= headers.length; j++) row.getCell(j).setCellStyle(cellStyles);
            }
        }
        
        public void createMovementsData(ArrayList<Movements> movements, String [] headers, XSSFWorkbook workbook, XSSFSheet sheet, XSSFCellStyle cellStyles) {
            XSSFFont fontIncome = workbook.createFont();
            fontIncome.setFontName("Times New Roman");
            fontIncome.setColor(IndexedColors.GREEN.getIndex());
            fontIncome.setBold(true);
            
            XSSFFont fontEgress = workbook.createFont();
            fontEgress.setFontName("Times New Roman");
            fontEgress.setColor(IndexedColors.RED.getIndex());
            fontEgress.setBold(true);
            
            XSSFFont boldFont = workbook.createFont();
            boldFont.setFontName("Times New Roman");
            boldFont.setBold(true);
            
            for(int i = 1; i <= movements.size(); i++) {
                XSSFRow row = sheet.createRow(i+1);
                
                XSSFCell incomeCell = row.createCell(1);
                XSSFCell egressCell = row.createCell(2);
                
                XSSFCellStyle incomeStyle = workbook.createCellStyle();
                incomeStyle.cloneStyleFrom(cellStyles);
                
                XSSFCellStyle egressStyle = workbook.createCellStyle();
                egressStyle.cloneStyleFrom(cellStyles);
                
                for(int j = 1; j < headers.length; j++) {
                    if(j == 1) {
                        incomeStyle.setFont(fontIncome);
                        incomeCell.setCellValue("+ $" + movements.get(i-1).getIncome());
                        incomeCell.setCellStyle(incomeStyle);
                    } else if (j == 2) {
                        egressStyle.setFont(fontEgress);
                        egressCell.setCellValue("- $" + movements.get(i-1).getEgress());
                        egressCell.setCellStyle(egressStyle);
                    }
                }
            }
            
            int totalIncome = 0, totalEgress = 0;
            for(int i = 0; i < movements.size(); i++) {
                totalIncome += movements.get(i).getIncome();
                totalEgress += movements.get(i).getEgress();
            }
            
            XSSFRow totalIncomeRow = sheet.getRow(2);
            XSSFRow totalEgressRow = sheet.getRow(3);
            XSSFRow totalProfitRow = sheet.getRow(4);
            
            XSSFCell totalIncomeCell = totalIncomeRow.createCell(Headers.movementsHeaders.length);
            XSSFCell totalEgressCell = totalEgressRow.createCell(Headers.movementsHeaders.length);
            XSSFCell totalProfitCell = totalProfitRow.createCell(Headers.movementsHeaders.length);
            
            totalIncomeCell.setCellStyle(cellStyles);
            totalEgressCell.setCellStyle(cellStyles);
            totalProfitCell.setCellStyle(cellStyles);
            
            XSSFRichTextString  totalIncomeStr = new XSSFRichTextString("TOTAL INCOME = +" + totalIncome);
            XSSFRichTextString  totalEgressStr = new XSSFRichTextString("TOTAL EGRESS = -" + totalEgress);
            XSSFRichTextString  totalProfitStr;
            
            totalIncomeStr.applyFont(0, totalIncomeStr.length(), fontIncome);
            totalEgressStr.applyFont(0, totalEgressStr.length(), fontEgress);
            
            totalIncomeCell.setCellValue(totalIncomeStr);
            totalEgressCell.setCellValue(totalEgressStr);
            
            if((totalIncome - totalEgress) >= 0) {
                totalProfitStr = new XSSFRichTextString("TOTAL PROFIT = +" + (totalIncome - totalEgress));
                totalProfitStr.applyFont(0, 13, boldFont);
                totalProfitStr.applyFont(14, totalIncomeStr.length(), fontIncome);
                totalProfitCell.setCellValue(totalProfitStr);
            } else {
                totalProfitStr = new XSSFRichTextString("TOTAL PROFIT = " + (totalIncome - totalEgress));
                totalProfitStr.applyFont(0, 13, boldFont);
                totalProfitStr.applyFont(14, totalIncomeStr.length(), fontEgress);
                totalProfitCell.setCellValue(totalProfitStr);
            }
        }

        
        public void autoSizeColumns(String [] headers, XSSFSheet sheet) {
            for(int i = 1; i <= headers.length; i++) {
                if(headers[i-1].length() != 0) sheet.autoSizeColumn(i);
            }
        }
    
}
