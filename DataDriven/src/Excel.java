
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;
import org.apache.commons.io.FileUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
public class Excel 
{
	public static void main(String[] args)
			throws InvalidFormatException, IOException,Throwable
{
try
{
	FileInputStream fls=new FileInputStream("D:\\parametrization req sw\\Test1.xlsx");
	Workbook wb=WorkbookFactory.create(fls);
	Sheet sh=wb.getSheet("Annu");
	int i=sh.getLastRowNum();
	for(int j=1; j<=i;j++)
	{	

	Row rw=sh.getRow(j);
	WebDriver Driver=new FirefoxDriver();
	String cellValue=rw.getCell(0).toString();
	/*//int m=(int)Float.parseFloat(cellValue);
	String emp=cellValue;
	String rehana="Rehana";
	if(emp.equals(rehana))
	{
		
	}*/

	Driver.get("http://www.proxio.com/crmls/user/"+cellValue+"");

	Driver.manage().window().maximize();
	Driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
	try
	{
		if(Driver.findElement(By.xpath("//*[@id='divAgentName']")).getText()!="")
		{
			String Property= Driver.findElement(By.xpath("//*[@id='divAgentName']")).getText();
			Driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
			System.out.println("Users Details Are :-"+Property);
			File scrFile = ((TakesScreenshot)Driver).getScreenshotAs(OutputType.FILE);
			FileUtils.copyFile(scrFile, new File("D:\\ScreenShot1\\"+Property+".png"));
			Driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		}
	}				
	 
	 catch(Exception e)
	  {
			String url = Driver.getCurrentUrl();
			System.out.println("This is an error page URL and the URL is::"+url);
			Driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
	  }
	 Driver.close();
}		 
			
}
catch (Exception e) 
{
	
}				
		
	
}
}
	
	
	

