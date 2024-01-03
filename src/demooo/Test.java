package demooo;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.time.Duration;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;

public class Test {

	public static void main(String[] args) throws IOException {
		WebDriver driver=new ChromeDriver();
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));
		driver.get("https://www.google.com/");
		WebElement wb=driver.findElement(By.name("btnK"));
		Actions act=new Actions(driver);
		act.contextClick(wb).perform();
		
		Select s=new Select(wb);
		s.selectByVisibleText("String visible text");
		
		FileInputStream fis=new FileInputStream("path of the Excell sheet");
		Workbook wbk=WorkbookFactory.create(fis);
		Sheet sh=wbk.getSheet("String Sheet name");
		Row rw=sh.getRow(2);
		Cell cel=rw.getCell(12);
		String val=cel.getStringCellValue();
		
		
		
	}

}
