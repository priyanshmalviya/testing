import java.io.File;  
import java.io.FileInputStream;
import java.util.Iterator;

import javax.jws.soap.SOAPBinding.Style;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;  
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class deloittetesting  
{  

public static void main(String[] args)   
{  
try  
{  
File file = new File("/Users/pmalviya/downloads/employee1.xlsx");   //creating a new file instance  
FileInputStream fis = new FileInputStream(file);   //obtaining bytes from the file  
//creating Workbook instance that refers to .xlsx file  
XSSFWorkbook wb = new XSSFWorkbook(fis);   
XSSFSheet sheet = wb.getSheetAt(0);     //creating a Sheet object to retrieve object  
Iterator<Row> itr = sheet.iterator();    //iterating over excel file  

sheet.getRow(0).createCell(7);
sheet.getRow(0).getCell(7).setCellValue("Number of days");



while (itr.hasNext())                 
{  
Row row = itr.next();  
Iterator<Cell> cellIterator = row.cellIterator();   //iterating over each column  
while (cellIterator.hasNext())   
{  
Cell cell = cellIterator.next();  
CellType type = cell.getCellTypeEnum();


switch (cell.getCellTypeEnum())               
{  
case STRING :    //field that represents string cell type 
	sheet.getRow(4).createCell(7).setCellFormula("SUM(E2:G2)");
	System.out.print(cell.getStringCellValue() + "\t\t");  
break;  
case NUMERIC:    //field that represents number cell type 		
System.out.print(cell.getNumericCellValue() + "\t\t\t\t");  
break;
	
default:  
}  
}  
System.out.println("");  
}  
}  
catch(Exception e)  
{  
e.printStackTrace();  
}  
}  
}  