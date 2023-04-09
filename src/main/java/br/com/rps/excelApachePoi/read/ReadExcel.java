package br.com.rps.excelApachePoi.read;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

public class ReadExcel {

    public void readExcelDataForSingleRow(String path) throws IOException {

        File f = new File(path);
        FileInputStream fis = new FileInputStream(f);
        Workbook workbook = WorkbookFactory.create(fis);
        Sheet sheet = workbook.getSheetAt(0);
        Row row = sheet.getRow(0);
        Cell cell = row.getCell(0);
        System.out.println(cell.getStringCellValue());
    }

    public void ReadExcelDataForEntireSheet(String path) throws IOException {

        File f = new File(path);
        FileInputStream fis = new FileInputStream(f);
        Workbook workbook = WorkbookFactory.create(fis);
        Iterator<Sheet> sheetIterator = workbook.sheetIterator();
        Sheet sheet = workbook.getSheetAt(0);
        Integer valorLinha = procurarCabecalho(sheet);

        while (sheetIterator.hasNext()) {
            sheet = sheetIterator.next();
            Iterator<Row> rowIterator = sheet.iterator();
            int currentRow = 1;
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    switch (cell.getCellType())
                    {
                        case BLANK:
//                            System.out.print("" + "\t");
                            break;
                        case BOOLEAN:
                            System.out.print(cell.getBooleanCellValue() + "\t");
                            break;
                        case ERROR:
                            System.out.print("error" + "\t");
                            break;
                        case NUMERIC:
                            System.out.print(cell.getNumericCellValue() + "\t");
                            break;

                        case FORMULA:
                            System.out.print(cell.getCellFormula() + "\t");
                            break;

                        case STRING:
                            System.out.print(cell.getStringCellValue() + "\t");
                            break;

                        case _NONE:
                            System.out.print("none" + "\t");
                            break;
                        default:
                            break;
                    }
                }
                System.out.println();
            }
        }
        System.out.println("fim");

    }

    public Integer procurarCabecalho(Sheet sheet){
        int startRow = 0;
        for (int i = 1; i <= 8; i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                Cell cell = row.getCell(0);
                if (cell != null && cell.getCellType() == CellType.STRING) {
                    String cellValue = cell.getStringCellValue().trim();
                    if (cellValue.equalsIgnoreCase("N°")) {
                        startRow = i +1 ;
                        System.out.println("Numero da linha é: " + startRow);
                        break;
                    }
                }
            }
        }

        return startRow;

    }




}



