package Excel_DataRead_Write;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel_DataWrite {
	public static void main(String[]args) throws IOException {
		
		
		
		InputStream  fis=null;
		Workbook  wbook=null;
		try {
			 fis=new FileInputStream("src\\main\\resources\\Book1.xlsx");
			  wbook=new XSSFWorkbook(fis);
		} catch (IOException e) {
			
			e.printStackTrace();
		}
		Sheet sheetobj=wbook.getSheet("Sheet1");
		
		int rowcount=sheetobj.getLastRowNum();   ///row count
		System.out.println(rowcount);
		Row  rowobj=sheetobj.getRow(1);
		
		short cellcount=rowobj.getLastCellNum();
		System.out.println(cellcount);
		Cell  cellobj=rowobj.getCell(2);
		
		String   datacount=cellobj.getStringCellValue();
		System.out.println(datacount);
	}

}
