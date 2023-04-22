 package utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
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
import org.openqa.selenium.chrome.ChromeOptions;
import org.testng.Assert;
import org.testng.AssertJUnit;
import org.testng.ITestResult;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import com.gargoylesoftware.htmlunit.ElementNotFoundException;

import old.APIClient;
import old.APIException;

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
 
public class UniversityUtility {
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
	  public static String TEST_RUN_ID   ;            
	   public static String TESTRAIL_USERNAME    ;      
	   public static String TESTRAIL_PASSWORD ;       
	    public static String RAILS_ENGINE_URL   ;       
	    public static final int TEST_CASE_PASSED_STATUS   = 1;
	    public static final int TEST_CASE_FAILED_STATUS   = 5;
	    public static String Logs = "";
	
	public void UniversityUtilityclass() throws IOException, InterruptedException {
		
		  
			
			
		
		
		//Open wakanow Site
		  String filePath = System.getProperty("user.dir");
		  System.setProperty("webdriver.chrome.driver",filePath +"//Drivers/chromedriver"); 
  	
		  
 	  FileInputStream fis = new FileInputStream(filePath+"/config/Testdata.xlsx");

	
	    
	    XSSFWorkbook workbook = new XSSFWorkbook(fis); 
		  
		  XSSFSheet sheets = workbook.getSheet("Environment");
		  XSSFSheet sheet = workbook.getSheet("Credentials");
		  XSSFSheet sheet2 = workbook.getSheet("Selenium");


		  objectrepo   = new Properties();
			File objfile = new File(System.getProperty("user.dir")+"//src//test//java//utils/objectrepo.properties");
			FileReader objrepo = new FileReader(objfile);
			objectrepo.load(objrepo);
		  
			 Row tEST_RUN_ID  = sheet2.getRow(1);
			  Cell cell20 = tEST_RUN_ID.getCell(1);
			  String RUN_ID   = cell20.getStringCellValue();
			  
			  Row tESTRAIL_USERNAME   = sheet2.getRow(2);
			  Cell cell21 = tESTRAIL_USERNAME.getCell(1);
			  String RAIL_USERNAME   = cell21.getStringCellValue();
			  
			  Row tESTRAIL_PASSWORD   = sheet2.getRow(3);
			  Cell cell22 = tESTRAIL_PASSWORD.getCell(1);
			  String RAIL_PASSWORD  = cell22.getStringCellValue();
			  
			  Row eNGINE_URL    = sheet2.getRow(4);
			  Cell cell23 = eNGINE_URL.getCell(1);
			  String ENGINE_URL   = cell23.getStringCellValue();
			  
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
		  
		  
		  Row username = sheet.getRow(13);
		  Cell cell1 = username.getCell(1);
		  String Username  = cell1.getStringCellValue();
		  
		  
		  Row password = sheet.getRow(14);
		  Cell cell7 = password.getCell(1);
		  String Password  = cell7.getStringCellValue();
		  
		   
			 
		  TEST_RUN_ID = RUN_ID;
		  TESTRAIL_USERNAME =RAIL_USERNAME ;
		   TESTRAIL_PASSWORD =  RAIL_PASSWORD  ;   
		     RAILS_ENGINE_URL =  ENGINE_URL ; 
		     
		     
		     System.setProperty("webdriver.chrome.silentOutput", "true");
		     java.util.logging.Logger.getLogger("org.openqa.selenium").setLevel(Level.OFF);
	    
		  
		  if (PLATFORM.equals("ALPHA")){
			 
			  driver = new ChromeDriver();	
			    log.info("Chrome browser is Selected ");
			  
			  baseUrl = STAGINGURL ;
			  
			  
				
			    
				  
				  
				  
				  // This code below will run test and not open the browser
				//  ChromeOptions options = new ChromeOptions();
				//  options.addArguments("--headless", "--window-size=1920,1200");
				 //   driver = new ChromeDriver(options);
				
				driver.manage().timeouts().implicitlyWait(40, TimeUnit.SECONDS);
				driver.manage().window().maximize();
				driver.manage().deleteAllCookies();
				driver.manage().window().maximize();
				driver.manage().deleteAllCookies();
				    driver.get(baseUrl);
			    log.info("Test Execution Started On  Alpha");
			 
			 
			    
					  
			}
			else if(PLATFORM.equals("PROD")){
				
				driver = new ChromeDriver();
				
			    
				  
				  
				  
				  // This code below will run test and not open the browser
				//  ChromeOptions options = new ChromeOptions();
				//  options.addArguments("--headless", "--window-size=1920,1200");
				 //   driver = new ChromeDriver(options);
				
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
			 // driver = new ChromeDriver();
					
				    
				  
				  
				  
				  // This code below will run test and not open the browser
		   ChromeOptions options = new ChromeOptions();
		  options.addArguments("--headless", "--window-size=1920,1200");
		   driver = new ChromeDriver(options);
				
				
				baseUrl = DEVURL;
				driver.manage().timeouts().implicitlyWait(40, TimeUnit.SECONDS);
				driver.manage().window().maximize();
				driver.manage().deleteAllCookies();
				driver.manage().window().maximize();
				driver.manage().deleteAllCookies();  
			    driver.get(baseUrl);
			    log.info("Test Execution Started On Dev");
			    
			   
			   
			}	

 

	   driver.findElement(By.xpath(objectrepo.getProperty("LoginEmailTextBox"))).sendKeys(Username);	 
	   Thread.sleep(2000);
	   driver.findElement(By.xpath(objectrepo.getProperty("LoginPasswordTextBox"))).sendKeys(Password);
	   Thread.sleep(2000);
	   
	  driver.findElement(By.xpath(objectrepo.getProperty("LoginButton"))).click();
	    log.info("Username , Password Entered and Sign In Button Clicked  ");
	   Thread.sleep(2000);
	
	    
	   Thread.sleep(2000);
		  driver.findElement(By.xpath(objectrepo.getProperty("MenuDropdown")));
		    log.info("University Login Successful");

		    Thread.sleep(2000);
}

	

	 public static void addResultForTestCase( String testCaseId, int status,
	            String error) throws IOException, APIException {

	        String testRunId = TEST_RUN_ID;
 
	        APIClient client = new APIClient(RAILS_ENGINE_URL);
	        client.setUser(TESTRAIL_USERNAME);
	        client.setPassword(TESTRAIL_PASSWORD);
	        Map data = new HashMap();
	        data.put("status_id", status);
         if (status ==TEST_CASE_PASSED_STATUS )
         {
 	 //  data.put("comment", "Test Executed - Status updated automatically from Selenium test automation.");
       	   data.put("comment", Logs);

        	 
         }
         else
         {
  data.put("comment", error);

         }
	        
	        client.sendPost("add_result_for_case/"+testRunId+"/"+testCaseId+"",data ); 

	  }

	        
	        
	        

 
  
  @AfterMethod //AfterMethod annotation - This method executes after every test execution
  public void screenShot(ITestResult result){

  driver.quit();
  }}