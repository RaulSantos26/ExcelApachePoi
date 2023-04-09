package br.com.rps.excelApachePoi.read;


import br.com.rps.excelApachePoi.entities.Product;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;


public class ProductController {


    public ArrayList<Product> getProducts(String path) throws IOException {
        ArrayList<Product> products = new ArrayList<>();
        int blankRowCount = 0;
        Workbook workbook;
        try {
            File f = new File(path);
            FileInputStream fis = new FileInputStream(f);
            workbook = WorkbookFactory.create(fis);
            Sheet sheet = workbook.getSheetAt(0);

            int startRow = 0;
            int rowNull = 0;
            for (int i = 0; i <= 15; i++) {
                Row row = sheet.getRow(i);
                if (row ==null){
                    rowNull++;
                }
                if (row != null) {
                    Cell cell = row.getCell(0);
                    if (cell != null && cell.getCellType() == CellType.STRING) {
                        String cellValue = cell.getStringCellValue().trim();
                        if (cellValue.equalsIgnoreCase("N°")) {
                            startRow = i+1;
                            break;
                        }
                    }
                }
            }

            Iterator<Row> rowIterator = sheet.rowIterator();
            int currentRow = 0;
            while (rowIterator.hasNext()) {
                Row row = (Row) rowIterator.next();
                if (isRowEmpty(row)) {
                    blankRowCount++;
                    continue;
                }
                if (currentRow < (startRow-rowNull)) {
                    // ignora as primeiras 7 linhas
                    currentRow++;
                    continue;
                }

                Cell nCell = row.getCell(0);
                Number n = (Number) getCellValue(nCell);

                Cell cdPrcCell = row.getCell(1);
                Number cdPrc = (Number) getCellValue(cdPrcCell);

                Cell cdPrfUndCell = row.getCell(2);
                Number prf = (Number) getCellValue(cdPrfUndCell);

                Cell cdEqpCell = row.getCell(3);
                String eqp = (String) getCellValue(cdEqpCell);

                Cell cdCliCell = row.getCell(4);
                Number cli = (Number) getCellValue(cdCliCell);

                Product produto = new Product(n, cdPrc, prf, eqp, cli);
                products.add(produto);


            }
        } catch (FileNotFoundException ex) {
            throw new RuntimeException(ex);
        } catch (IOException ex) {
            throw new RuntimeException(ex);
        }
        workbook.close();
        System.out.println(products);
        return products;
    }

    private static Object getCellValue(Cell cell) {
        if (cell == null) {
            return null;
        }
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (org.apache.poi.ss.usermodel.DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue();
                } else {
                    double value = cell.getNumericCellValue();
                    if (value == (int) value) {
                        return (int) value;
                    } else {
                        return value;
                    }
                }
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case BLANK:
            case _NONE:
                return null;
            default:
                return null;
        }
    }
    private static boolean isRowEmpty(Row row) {
        if (row == null) {
            return true;
        }

        Iterator<Cell> cellIterator = row.cellIterator();
        while (cellIterator.hasNext()) {
            Cell cell = cellIterator.next();
            Object cellValue = getCellValue(cell);
            if (cellValue != null) {
                return false;
            }
        }

        return true;
    }
}



