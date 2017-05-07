package testImport;

import java.io.*;
import java.util.Iterator;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;

public class Test1 {

	public static void main(String[] args) throws Exception {

		// path for the file
		 String data = "D:/TestSoPro/testFile.xlsx";

		String path = "D:/TestSoPro/";
		String nameFile = "testFile.xlsx";

		File file = new File(path + nameFile);
		FileInputStream fis = new FileInputStream(file);

		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet spreadsheet = workbook.getSheetAt(1);

		//Row r = spreadsheet.getRow(0);
		//int maxCell = r.getLastCellNum();

		// System.out.println("The number of columns is: " + maxCell);
		int nrRows = spreadsheet.getPhysicalNumberOfRows();
		int lastRow = spreadsheet.getLastRowNum();
		int firstRow = spreadsheet.getFirstRowNum();
		System.out.println("THe number of rows is: " + nrRows);
		System.out.println("The first row is: " + firstRow + " and the last row is: " + lastRow);

	}

}
