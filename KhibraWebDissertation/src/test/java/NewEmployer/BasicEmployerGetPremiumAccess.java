 package NewEmployer;

import org.testng.annotations.Test;
import org.testng.AssertJUnit;
import org.testng.annotations.Test;
import org.testng.Assert;
import org.testng.annotations.Test;
import org.testng.AssertJUnit;
import org.testng.annotations.Test;
import org.testng.Assert;
import org.testng.annotations.Test;
import org.testng.Assert;
import org.testng.annotations.Test;
import org.testng.AssertJUnit;
import org.testng.annotations.Test;
import org.testng.Assert;
import org.testng.annotations.Test;
import org.testng.AssertJUnit;
import org.testng.annotations.Test;
import org.testng.Assert;
import org.testng.annotations.Test;
import org.testng.AssertJUnit;
import org.testng.annotations.Test;
import org.testng.Assert;
import org.testng.annotations.Test;
import org.testng.AssertJUnit;
import org.testng.annotations.Test;
import org.testng.Assert;
import org.testng.annotations.Test;
import org.testng.AssertJUnit;
import org.testng.annotations.Test;
import org.testng.annotations.Test;
import org.testng.annotations.Test;
import org.testng.annotations.Test;
import org.testng.annotations.Test;
import static org.testng.AssertJUnit.assertFalse;
import static org.testng.AssertJUnit.assertTrue;
import org.testng.annotations.Test;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Date;
import java.util.Iterator;
import java.util.Set;
import java.util.concurrent.TimeUnit;
import java.util.regex.Pattern;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.NoAlertPresentException;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;

import static org.hamcrest.CoreMatchers.*;
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
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import com.gargoylesoftware.htmlunit.ElementNotFoundException;

import utils.BasicEmployerUtility;
import utils.EmployerUtility;
import utils.log;
import utils.utility;

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
 
public class BasicEmployerGetPremiumAccess extends BasicEmployerUtility{
	 
		
	@BeforeMethod
    public void setUp1() throws Exception {
	  log.info("START=BasicEmployerGetPremiumAccess--------------------------------------------------------------");

	  BasicEmployerUtilityclass();
 	
}
	
	 
	@SuppressWarnings("rawtypes")
	@Test
	
	public void testBasicEmployerGetPremiumAccess() throws Exception {
	     
		  try {
			  Thread.sleep(2000);
			  
			    
				  
   driver.findElement(By.xpath(objectrepo.getProperty("EmpDashboardMenu"))).click();	 
		  log.info("Basic Emp Dashboard Menu  Clicked");
		  Thread.sleep(2000);
		   driver.findElement(By.xpath(objectrepo.getProperty("DashboardHeader")));	 

		   log.info("Basic Employer Dashboard Menu  Opened");
		   
		   
		   driver.findElement(By.xpath(objectrepo.getProperty("GetPremiumAccessButton"))).click();	
		   log.info("Get Premium Access Button Clicked ");
		   driver.findElement(By.xpath(objectrepo.getProperty("BasicText")));	
		   log.info("Account Settings Page Opened ");
		   driver.findElement(By.xpath(objectrepo.getProperty("AccountBasicText")));	
		   driver.findElement(By.xpath(objectrepo.getProperty("AccountFeaturesText")));	
		   driver.findElement(By.xpath(objectrepo.getProperty("1MonthPlan")));	
		   
		   /*
		   driver.findElement(By.xpath(objectrepo.getProperty("3MonthsPlan")));	
		   driver.findElement(By.xpath(objectrepo.getProperty("6MonthsPlan")));	
		   driver.findElement(By.xpath(objectrepo.getProperty("12MonthsPlan")));	
		   driver.findElement(By.xpath(objectrepo.getProperty("PlanAmount")));	
		   log.info("Account Features, Plans and Price Validated ");
		   driver.findElement(By.xpath(objectrepo.getProperty("1MonthPlan"))).click();
		   Thread.sleep(2000);
		   log.info("Plan Selected ");
		   driver.findElement(By.xpath(objectrepo.getProperty("BuyPremiumButton"))).click();
		   Thread.sleep(2000);
		   log.info("Buy Premium Button Clicked");
 		    
  		   
		   log.info("Payment Gateway is Opened Successfully");
		  
		   */
		   Logs =   " log.info(\"Basic Emp Dashboard Menu  Clicked\");\r\n"
					+ "		  log.info(\"Basic Employer Dashboard Menu  Opened \");\r\n"
					+ "		  log.info(\"Get Premium Access Button Clicked \");\r\n"
					+ "		  log.info(\"Account Settings Page Opened \");\r\n"

					+ "		  log.info(\"Account Features, Plans and Price Validated  \");\r\n"

					+ "		  log.info(\"Plan Selected  \");\r\n"
					+ "		  log.info(\"Buy Premium Button Clicked \");\r\n"
					+ "		  log.info(\"Payment Gateway is Opened Successfully \");\r\n"

	;

			    log.info("End TEST-------------------------- BasicEmployerGetPremiumAccess------------------------");     
	    
			    log.info("END=PASSED");
			    
			  
			    BasicEmployerGetPremiumAccess.addResultForTestCase("1317", TEST_CASE_PASSED_STATUS, ""); 	    
			     
		  } 
		  
		  
		  catch (Exception e) {
			  log.info("END=FAILED");
			  
 
			  BasicEmployerGetPremiumAccess.addResultForTestCase("1317", TEST_CASE_FAILED_STATUS,  e.getMessage()); 	    
		   AssertJUnit.fail(e.getMessage());
								 
		 	}
		 }
  
  
  
  @AfterMethod //AfterMethod annotation - This method executes after every test execution
  public void screenShot(ITestResult result){
  //using ITestResult.FAILURE is equals to result.getStatus then it enter into if condition
  if(ITestResult.FAILURE==result.getStatus()){
  try{
  // To create reference of TakesScreenshot
  TakesScreenshot screenshot=(TakesScreenshot)driver;
  // Call method to capture screenshot
  File src=screenshot.getScreenshotAs(OutputType.FILE);
  // Copy files to specific location 
  // result.getName() will return name of test case so that screenshot name will be same as test case name
  FileUtils.copyFile(src, new File(System.getProperty("user.dir")+"/screenshots/"+result.getName()+new SimpleDateFormat("yyyyMMddhhmm").format(new Date())+".jpg"));

  System.out.println("Successfully captured a screenshot");
  }catch (Exception e){
  System.out.println("Exception while taking screenshot "+e.getMessage());
  } 
  }
  driver.quit();
  }
  
  private boolean isElementPresent(By by) {
	    try {
	      driver.findElement(by);
	      return true;
	    } catch (NoSuchElementException e) {
	      return false;
	    }
	  }
  
  private boolean isAlertPresent() {
	    try {
	      driver.switchTo().alert();
	      return true;
	    } catch (NoAlertPresentException e) {
	      return false;
	    }
	  }

	   
 
  
}
