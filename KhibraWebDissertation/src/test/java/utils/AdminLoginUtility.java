 package utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.util.Date;
import java.util.Properties;
import java.util.concurrent.TimeUnit;
import java.util.logging.Level;
import java.util.regex.Pattern;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.NoAlertPresentException;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;

import static org.hamcrest.CoreMatchers.*;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;
import static org.testng.Assert.assertEquals;

import org.apache.commons.io.FileUtils;
import org.apache.log4j.BasicConfigurator;
import org.apache.log4j.Logger;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.Assert;
import org.testng.AssertJUnit;
import org.testng.ITestResult;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import com.gargoylesoftware.htmlunit.ElementNotFoundException;

import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import java.io.FileNotFoundException;
 
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class AdminLoginUtility {
	public WebDriver driver;
	private String baseUrl;
	  private boolean acceptNextAlert = true;
	  private StringBuffer verificationErrors = new StringBuffer();
	  private static Workbook wb;
	  private static Sheet sh;
	  private static FileInputStream fis; 
	  private static FileOutputStream fos; 
	  private static Row row;
	  private static Cell cell;
	  protected Properties objectrepo;
	  protected Properties config;	
	
	
	public void AdminLoginUtility() throws IOException, InterruptedException {
		
		  
			
			
		
		
 
		  String filePath = System.getProperty("user.dir");
		
		 
		  
 	  FileInputStream fis = new FileInputStream(filePath+"/config/Testdata.xlsx");

	
	    
	    XSSFWorkbook workbook = new XSSFWorkbook(fis); 
		  
		  XSSFSheet sheets = workbook.getSheet("Environment");
		  XSSFSheet sheet = workbook.getSheet("Credentials");


		  objectrepo   = new Properties();
			File objfile = new File(System.getProperty("user.dir")+"//src//test//java//utils/objectrepo.properties");
			FileReader objrepo = new FileReader(objfile);
			objectrepo.load(objrepo);
		  
		  
		  Row plaform = sheets.getRow(1);
		  Cell cell11 = plaform.getCell(1);
		  String PLATFORM  = cell11.getStringCellValue();
		  
		  
		  
		  Row STAGEURL= sheets.getRow(2);
		  Cell cell5 = STAGEURL.getCell(1);
		  String STAGINGURL  = cell5.getStringCellValue();
		  
		  Row PRODUCTIONURL= sheets.getRow(3);
		  Cell cell4 = PRODUCTIONURL.getCell(1);
		  String PRODURL  = cell4.getStringCellValue();
		  
		  Row dEVURL= sheets.getRow(4);
		  Cell cell2 = dEVURL.getCell(1);
		  String DEVURL  = cell2.getStringCellValue();
		  
		  
		  
		  Row username = sheet.getRow(9);
		  Cell cell1 = username.getCell(1);
		  String Username  = cell1.getStringCellValue();
		  
		  
		  Row password = sheet.getRow(10);
		  Cell cell7 = password.getCell(1);
		  String Password  = cell7.getStringCellValue();
		  
		  Row browser = sheets.getRow(5);
		  Cell cell8 = browser.getCell(1);
		  String Browser  = cell8.getStringCellValue();
		  
		  System.setProperty("webdriver.chrome.silentOutput", "true");
		  java.util.logging.Logger.getLogger("org.openqa.selenium").setLevel(Level.OFF);
		   
		  {
			  
			  if (Browser.equals("CHROME")){
				  
				  System.setProperty("webdriver.chrome.driver",filePath +"//Drivers/chromedriver"); 
					driver = new ChromeDriver();	
				    log.info("Chrome browser is Selected ");
		 			  
				}
				else if(Browser.equals("FIREFOX")){
			 
				    log.info("Chrome browser is Selected ");   
				}	
			  		
		  }
		  
		  
		  
		  
		  
	    
		  
		  if (PLATFORM.equals("ALPHA")){
			  baseUrl = STAGINGURL ;
				driver.manage().timeouts().implicitlyWait(40, TimeUnit.SECONDS);
				driver.manage().window().maximize();
				driver.manage().deleteAllCookies();
				driver.manage().window().maximize();
				driver.manage().deleteAllCookies();
				    driver.get(baseUrl);
			    log.info("Test Execution Started On  Alpha");
			 
			   
			    
					  
			}
			else if(PLATFORM.equals("PROD")){
				baseUrl = PRODURL;
				driver.manage().timeouts().implicitlyWait(40, TimeUnit.SECONDS);
				driver.manage().window().maximize();
				driver.manage().deleteAllCookies();
				driver.manage().window().maximize();
				driver.manage().deleteAllCookies();  
			    driver.get(baseUrl);
			    log.info("Test Execution Started On Production");
			   
			}	
		 
			else if(PLATFORM.equals("DEV")){
				baseUrl = DEVURL;
				driver.manage().timeouts().implicitlyWait(40, TimeUnit.SECONDS);
				driver.manage().window().maximize();
				driver.manage().deleteAllCookies();
				driver.manage().window().maximize();
				driver.manage().deleteAllCookies();  
			    driver.get(baseUrl);
			    log.info("Test Execution Started On Dev");
			   
			}	

 

		 
  
		  driver.findElement(By.xpath(objectrepo.getProperty("LoginButton"))).click();	   
		  Thread.sleep(2000);
		    log.info("Login Button Clicked ");
			 
		   

	   driver.findElement(By.xpath(objectrepo.getProperty("LogEmailTextBox"))).sendKeys(Username);	 
	   Thread.sleep(2000);
	   driver.findElement(By.xpath(objectrepo.getProperty("LogPasswordTextbox"))).sendKeys(Password);
	   Thread.sleep(2000);
	   
	  driver.findElement(By.xpath(objectrepo.getProperty("SignInButton"))).click();
	    log.info("Username , Password  and Sign In Button Entered  ");
	   Thread.sleep(2000);
	

	   
		  driver.findElement(By.xpath(objectrepo.getProperty("UsernameDropdown")));
		    log.info("Login Successful");


}

 
  @AfterMethod //AfterMethod annotation - This method executes after every test execution
  public void screenShot(ITestResult result){

  driver.quit();
  }}