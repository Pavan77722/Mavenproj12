package mday1;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class MavenXmlSheet {

	public static void main(String[] args) throws IOException {
		FileInputStream file=new FileInputStream("D:\\java\\FirstMavenProj\\xlsheet\\Emptbl.xlsx");
		XSSFWorkbook wb=new XSSFWorkbook(file);
		XSSFSheet sheet=wb.getSheet("Sheet1");
		int rows=sheet.getLastRowNum();
		int columns=sheet.getRow(rows ).getLastCellNum();
		
		System.out.println(rows);
		System.out.println(columns);
		
		for(int i=1;i<=rows;i++) {
			XSSFRow row=sheet.getRow(i);
			
			for(int c=0;c<columns;c++) {
				String value=row.getCell(c).toString();
				System.out.print(value);
			}
			System.out.println();
		}

	}

}
