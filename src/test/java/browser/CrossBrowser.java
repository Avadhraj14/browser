package browser;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.testng.ITestResult;
import org.testng.annotations.AfterClass;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Parameters;
import org.testng.annotations.Test;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

public class CrossBrowser {
	
	public WebDriver driver;

	public ExtentTest test;
	public ExtentReports report;
	public ExtentTest logger;
	
	 @BeforeTest
	    public void init()
	    {
	        report = new ExtentReports(System.getProperty("user.dir") + "/test-output/ExtentScreenshot.html", true);
	    }
	
	  @Parameters("browser")

	  @BeforeClass
	  public void beforeTest(String browser)
	  {
		  report = new ExtentReports(System.getProperty("user.dir")+"\\ExtentReportResults.html");
			test = report.startTest("CrossBrowser");
		  if(browser.equalsIgnoreCase("firefox")) 
		  {
			  System.setProperty("webdriver.gecko.driver", "C:\\Selenium\\Browser\\geckodriver.exe");
			  driver = new FirefoxDriver();	  
			  driver.manage().window().maximize();

			  test.log(LogStatus.PASS, "Open Firefox Browser");
		  }
		  else if (browser.equalsIgnoreCase("chrome")) 
		  { 
			  System.setProperty("webdriver.chrome.driver", "C:\\Selenium\\Browser\\chromedriver.exe");
			  driver = new ChromeDriver();
			  driver.manage().window().maximize();
			  test.log(LogStatus.PASS, "Open Chrome Browser");
		  } 
	  
		  else if (browser.equalsIgnoreCase("edge")) 
		  { 
			  System.setProperty("webdriver.edge.driver", "C:\\Selenium\\Browser\\msedgedriver.exe");
			  driver = new EdgeDriver();
			  driver.manage().window().maximize();
			  test.log(LogStatus.PASS, "Open Edge Browser");
					  
		  } 

	  driver.get("https://web.vepaar.com/?utm_campaign=header_cta&utm_medium=all&utm_source=vepaar_website#/login"); 
	
	  
	  
	  
	  }
	  
	  
	 @Test
	 @Parameters("Exceldata")

	public void login() throws InterruptedException, IOException 
	  {
		    XSSFWorkbook workbook;
		    XSSFSheet sheet;
		    XSSFCell cell;

		    File src=new File("Book.xlsx");
		    FileInputStream finput = new FileInputStream(src);
		    workbook = new XSSFWorkbook(finput);
		    sheet= workbook.getSheetAt(0);
		      
		     for(int i=1; i<=sheet.getLastRowNum(); i++)
		     {
		         cell = sheet.getRow(i).getCell(0);
		         cell.setCellType(CellType.STRING);
		        System.out.println(cell.getStringCellValue());
		        WebElement username= driver.findElement(By.xpath("//input[@type='text']"));
		        username.sendKeys(cell.getStringCellValue());
		     
		         cell = sheet.getRow(i).getCell(1);
		         cell.setCellType(CellType.STRING);
		        System.out.println(cell.getStringCellValue());
		         driver.findElement(By.xpath("//input[@type='password']")).sendKeys(cell.getStringCellValue());
		         driver.findElement(By.xpath("//button[@type='submit']")).click();    
		         
		         Thread.sleep(2000);
		         
		         ((JavascriptExecutor) driver).executeScript("window.open()");
		         ArrayList<String> tabs = new ArrayList<String>(driver.getWindowHandles());
		         driver.switchTo().window(tabs.get(i));
		         driver.get("https://web.vepaar.com/?utm_campaign=header_cta&utm_medium=all&utm_source=vepaar_website#/login");
		         test.log(LogStatus.PASS, "get data from excel sheet"); 
		        }
		} 
	 
	 public static String capture(WebDriver driver,String screenShotName) throws IOException
	    {
	        TakesScreenshot ts = (TakesScreenshot)driver;
	        File source = ts.getScreenshotAs(OutputType.FILE);
	        String dest = System.getProperty("user.dir") +"\\ErrorScreenshots\\"+screenShotName+".png";
	        File destination = new File(dest);
	        FileUtils.copyFile(source, destination);        
	                     
	        return dest;
	    }
	  
	 @AfterMethod
	    public void getResult(ITestResult result) throws IOException
	    {
		 if(result.getStatus() == ITestResult.FAILURE)
	        {
	        
	            String screenShotPath = CrossBrowser.capture(driver, "screenShotName");
	            test.log(LogStatus.FAIL, result.getThrowable());
	            test.log(LogStatus.FAIL, "Snapshot below: " + test.addScreenCapture(screenShotPath));
	        } 
	        report.endTest(test);
	    }
	  
	  @AfterClass
	  public void afterTest() 
	  {
		  report.endTest(test);
		  report.flush();
			driver.quit();
	  }

}
