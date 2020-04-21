package firstTestNGPackage;

import org.testng.ITestResult;
import org.testng.SkipException;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.remote.session.CapabilitiesFilter;
import org.testng.asserts.*;
import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;


import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import junit.framework.Assert;


public class FirstTestNGfile {

	ExtentReports extent;
	ExtentTest logger;
	
@BeforeTest
public void startTest() {
	extent = new ExtentReports(System.getProperty("user.dir")+"/Results/Reports.html", true);
	extent.loadConfig(new File(System.getProperty("user.dir")+"\\extent-config.xml"));
}
  @Test
  public void passTC() throws IOException {
	  
	 
	  logger = extent.startTest("PassTc");
	  
	  System.out.println("Inside the testNg, yay");
	  
	  
	  System.setProperty("webdriver.gecko.driver", "D:\\Java Software\\gecko driver\\geckodriver.exe");
	  WebDriver driver = new FirefoxDriver();
	  System.out.println("still working");
	  String baseurl = "https://www.w3schools.com/sql/sql_syntax.asp";
	  driver.get(baseurl);
	 
   
	  List<WebElement> rows = driver.findElements(By.xpath("//table[@class='w3-table-all notranslate']//tr"));
	  int noOfRows = rows.size();
	  String  beforeXpath = "//table[@class='w3-table-all notranslate']//tr[]//th[]";
	  System.out.println(noOfRows);
	  
	for(int i = 2; i <= noOfRows;i++) {
		String x = driver.findElement(By.xpath("//table[@class='w3-table-all notranslate']//tr["+i+"]//td[1]")).getText();
		System.out.println(x+ " x");
		writeToExcel("D:\\LearnGIT\\Excel.xlsx", "Sheet1", i, 1, x);
		
	}
	
	for(int i = 3; i <= noOfRows;i++) {
		String x = driver.findElement(By.xpath("//table[@class='w3-table-all notranslate']//tr["+i+"]//td[1]")).getText();
		System.out.println(x+ " x");
		writeToExcel("D:\\LearnGIT\\Excel.xlsx", "Sheet1", i, 1, x);
		
	}
	  
	
	for(int i = 2; i <= noOfRows;i++) {
		String x = driver.findElement(By.xpath("//table[@class='w3-table-all notranslate']//tr["+i+"]//td[1]")).getText();
		System.out.println(x+ " x");
		writeToExcel("D:\\LearnGIT\\Excel.xlsx", "Sheet1", i, 1, x);
		
	}
	  
	
	for(int i = 2; i <= noOfRows;i++) {
		String x = driver.findElement(By.xpath("//table[@class='w3-table-all notranslate']//tr["+i+"]//td[1]")).getText();
		System.out.println(x+ " x");
		writeToExcel("D:\\LearnGIT\\Excel.xlsx", "Sheet1", i, 1, x);
		
	}
	  
	
	for(int i = 2; i <= noOfRows;i++) {
		String x = driver.findElement(By.xpath("//table[@class='w3-table-all notranslate']//tr["+i+"]//td[1]")).getText();
		System.out.println(x+ " x");
		writeToExcel("D:\\LearnGIT\\Excel.xlsx", "Sheet1", i, 1, x);
		
	}
	  
	  
	  
	  
	  //table[@class='w3-table-all notranslate']//tr
	  
//	  driver.quit();
	  logger.log(LogStatus.PASS, "TC passed");
  }
  
  public void writeToExcel(String excelPath, String sheetname, int rows, int col, String value) throws IOException {
	  FileInputStream file = new FileInputStream(new File(excelPath));
	  FileOutputStream fos = null;
      XSSFWorkbook workbook = new XSSFWorkbook(file);      
      XSSFSheet sheet = workbook.getSheet(sheetname);
      XSSFRow row  = null;
      XSSFCell cell = null;
      
      row  = sheet.getRow(rows);
      if(row ==null) 
    	  row = sheet.createRow(rows);
     cell = row.getCell(col)	 ;
     if(cell == null)
    	 cell= row.createCell(col);
      
     cell.setCellValue(value);
     fos = new FileOutputStream(excelPath);
     workbook.write(fos);
	  fos.close();
  }
  
  


  @AfterMethod
  public void getReslt(ITestResult result) {
	  
	  
	  if(result.getStatus()==ITestResult.FAILURE) {
		  
		  logger.log(LogStatus.FAIL, "Test case failed is "+result.getName());
	  }
	  else if (result.getStatus()==ITestResult.SKIP) {
		  logger.log(LogStatus.SKIP, "Test case skipped is "+result.getName());
		
	}
	 extent.endTest(logger);
	  
  }
  
  @AfterTest
  public void endReport() {
	  extent.flush();
	  extent.close();
  }
}
