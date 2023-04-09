package br.com.rps.excelApachePoi.write;


import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class WriteExcel {

    public void writeSingleCellData(String filePath) throws IOException {

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("firstSheet");
        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);
        cell.setCellValue("first Cell");

        File f = new File(filePath);
        FileOutputStream fos = new FileOutputStream(f);
        workbook.write(fos);

        fos.close();
        workbook.close();
    }

    public void writeMultipleCellData(String filePath) throws IOException {

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("firstSheet");
        int[][] dataArray = getRandomDataArray(5,6);
        for (int i = 0 ; i< dataArray.length; i++){
            Row row = sheet.createRow(i);
            for (int j =0; j< dataArray[i].length; j++){
                Cell cell = row.createCell(j);
                cell.setCellValue(dataArray[i][j]);
            }
        }
        File f = new File(filePath);
        FileOutputStream fos = new FileOutputStream(f);
        workbook.write(fos);
        fos.close();
        workbook.close();
    }

    private int[][] getRandomDataArray(int row, int col) {
        int[][] dataArray = new int[row][col];
        for(int i=0; i<dataArray.length; i++){
            for (int j=0; j<dataArray[i].length; j++){
                dataArray[i][j] = (int)(Math.random()*1000);
            }
        }
        return dataArray;
    }
}
