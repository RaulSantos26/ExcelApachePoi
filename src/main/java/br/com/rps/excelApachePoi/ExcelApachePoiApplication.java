package br.com.rps.excelApachePoi;

import br.com.rps.excelApachePoi.read.ProductController;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.IOException;

@SpringBootApplication
public class ExcelApachePoiApplication {

	public static void main(String[] args) throws IOException {
		SpringApplication.run(ExcelApachePoiApplication.class, args);
		String path ="C:\\Temp\\planilhasExcel\\prc.xlsx";
		/*
		Criar uma planilhaS
		WriteExcel write = new WriteExcel();
//		write.writeSingleCellData(path);
		write.writeMultipleCellData(path);
		 */

//		ReadExcel read = new ReadExcel();
//		read.readExcelDataForSingleRow(path);
//		read.ReadExcelDataForEntireSheet(path);

		ProductController productController = new ProductController();
		productController.getProducts(path);


	}



}
