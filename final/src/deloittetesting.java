import java.io.File;
import java.io.FileInputStream;

import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;

import org.apache.poi.ss.usermodel.CellType;

import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class deloittetesting {

	public static void main(String[] args) {
		try {
			File file = new File("/Users/pmalviya/downloads/employee1.xlsx"); // creating a new file instance
			FileInputStream fis = new FileInputStream(file); // obtaining bytes from the file

//creating Workbook instance that refers to .xlsx file  
			XSSFWorkbook wb = new XSSFWorkbook(fis);

			XSSFSheet sheet = wb.getSheetAt(0); // creating a Sheet object to retrieve object
			Iterator<Row> itr = sheet.iterator(); // iterating over excel file

			sheet.getRow(0).createCell(7);
			sheet.getRow(0).getCell(7).setCellValue("Number of days"); // adding new column

			String date2 = "14-02-2020";

			while (itr.hasNext()) {
				Row row = itr.next();
				Iterator<Cell> cellIterator = row.cellIterator(); // iterating over each column
				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();
					CellType type = cell.getCellTypeEnum();

					switch (cell.getCellTypeEnum()) {
					case STRING: // field that represents string cell type

						System.out.print(cell.getStringCellValue() + "\t\t");
						break;
					case NUMERIC: // field that represents number cell type

						if (DateUtil.isCellDateFormatted(cell)) {
							SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy"); // Converting date format
																								// from excel
							System.out.print(dateFormat.format(cell.getDateCellValue()) + "\t\t");
						} else {
							System.out.print(cell.getNumericCellValue() + "\t\t\t\t");
						}
						break;

					default:
					}
				}
				System.out.println("");
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
