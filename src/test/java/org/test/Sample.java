package org.test;

import java.io.File;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;

public class Sample {
	
	public static void main(String[] args) throws IOException {
		
		System.setProperty("webdriver.chrome.driver","C:\\Users\\Suvetha P S\\eclipse-workspace\\SampleProject\\Drivers\\chromedriver.exe");
		WebDriver driver = new ChromeDriver();
		driver.get("http://demo.automationtesting.in/Register.html");
		driver.manage().window().maximize();
		WebElement ddncountry = driver.findElement(By.id("countries"));
		Select select= new Select(ddncountry);
		List<WebElement> allOptions = select.getOptions();
		System.out.println(allOptions);
		File file= new File("C:\\Users\\Suvetha P S\\eclipse-workspace\\SampleProject\\Excel\\sample123.xlsx");
		Workbook workbook = new XSSFWorkbook();
		Sheet sheet = workbook.createSheet("Sample");
		for (int i = 0; i < allOptions.size(); i++) {
			WebElement element = allOptions.get(i);
			String name = element.getText();
			System.out.println(name);
			Row row = sheet.createRow(i);
			Cell cell = row.createCell(0);
			cell.setCellValue(name);
			
	}
		FileOutputStream fileOutputStream = new FileOutputStream(file);
			workbook.write(fileOutputStream);
				System.out.println("done!!!!");
				
			}

}
