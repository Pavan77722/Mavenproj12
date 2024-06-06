package day3;

import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class LinksinExcel {

	public static WebDriver driver;
	public static void main(String[] args) throws IOException, InterruptedException {
		
		driver=new ChromeDriver();
		
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));
		driver.get("https://www.amazon.com/");
		Thread.sleep(10000);
		driver.manage().window().maximize();
		driver.findElement(By.id("searchDropdownBox")).click();
		driver.findElement(By.xpath("//select[@id='searchDropdownBox']/option[6]")).click();
		driver.findElement(By.id("nav-search-submit-button")).click();
		
		List<WebElement>links=driver.findElements(By.tagName("a"));
		
		int size=links.size();
		FileOutputStream fo=new FileOutputStream("D:\\java\\FirstMavenProj\\ExcelPrac11\\Prac_ExcelLinks.xlsx");
		XSSFWorkbook wb=new XSSFWorkbook();
		XSSFSheet sheet=wb.createSheet("Sheet1");
		
		for(int i=0;i<size;i++) {
			
			String values=links.get(i).getText();
			
			XSSFRow row=sheet.getRow(i);
			
			if(row==null) {
				row=sheet.createRow(i);
				
				XSSFCell cell=row.createCell(0);
				cell.setCellValue(values);
			}
			
			
			
			wb.write(fo);
			System.out.println(values);

	}

}
} 

