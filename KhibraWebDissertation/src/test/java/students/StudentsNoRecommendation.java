 package students;

import org.testng.annotations.Test;
import org.testng.Assert;
import org.testng.AssertJUnit;
import org.testng.annotations.Test;
import org.testng.Assert;
import org.testng.AssertJUnit;
import org.testng.annotations.Test;
import org.testng.Assert;
import org.testng.AssertJUnit;
import org.testng.annotations.Test;
import org.testng.Assert;
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

import utils.NoCVLoginUtility;
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
 
public class StudentsNoRecommendation extends NoCVLoginUtility{
	 
		
	@BeforeMethod
    public void setUp1() throws Exception {
	  log.info("START=StudentsNoRecommendation--------------------------------------------------------------");

	  NoCVLoginUtilityclass();
    
    	
    		
}
	
	
	  @SuppressWarnings("rawtypes")
	@Test
	
	public void testStudentsNoRecommendation() throws Exception {
		  try {
		   
			  
			  driver.findElement(By.xpath(objectrepo.getProperty("MenuDropdown"))).click();	 
			  log.info("Menu Dropdown Clicked");
			  driver.findElement(By.xpath(objectrepo.getProperty("DashboardMenu"))).click();	 
			  log.info("Dashboard Menu Selected");
			  Thread.sleep(2000);
			  driver.findElement(By.xpath(objectrepo.getProperty("NoJobApplied")));	 

			  driver.findElement(By.xpath(objectrepo.getProperty("EmptyJobImage")));	 
			 
			  log.info("All Empty Jobs  Images and texts Dispalyed");

			  
   if (driver.findElements(By.xpath(objectrepo.getProperty("RecommendedJobText"))).size()>0) {
			  
				   TakesScreenshot screenshot=(TakesScreenshot)driver;
				   		  // Call method to capture screenshot
			   File src=screenshot.getScreenshotAs(OutputType.FILE);
		 	  // Copy files to specific location 
			   // result.getName() will return name of test case so that screenshot name will be same as test case name
	      	   FileUtils.copyFile(src, new File(System.getProperty("user.dir")+"\\screenshots\\"+new SimpleDateFormat("yyyyMMddhhmm").format(new Date())+".jpg"));
				   		 
				    Assert.fail("Recommendations  was found for a New User on the Homepage  ");
				   	  
				      }
				   	  else 
				   		  
				   	  {
				   		  log.info("Recommendations not Available for New Users as Expected "); 
				   	  }
			  
   driver.findElement(By.xpath(objectrepo.getProperty("MenuDropdown"))).click();	 
	  log.info("Menu Dropdown Clicked");
	  driver.findElement(By.xpath(objectrepo.getProperty("JobsMenus"))).click();	 
	  log.info("Jobs Menu Selected");
	  Thread.sleep(2000);
	  driver.findElement(By.xpath(objectrepo.getProperty("JobsHeader")));	 
			  
			  
	  
	  
if (driver.findElements(By.xpath(objectrepo.getProperty("RecommendedLabel1"))).size()>0) {
	  
		   TakesScreenshot screenshot=(TakesScreenshot)driver;
		   		  // Call method to capture screenshot
	   File src=screenshot.getScreenshotAs(OutputType.FILE);
 	  // Copy files to specific location 
	   // result.getName() will return name of test case so that screenshot name will be same as test case name
  	   FileUtils.copyFile(src, new File(System.getProperty("user.dir")+"\\screenshots\\"+new SimpleDateFormat("yyyyMMddhhmm").format(new Date())+".jpg"));
		   		 
		    Assert.fail("Recommendations  was found for a New User on the Jobs Page  ");
 
		      }
		   	  else 
		   		  
		   	  {
		   		  log.info("Recommendations not Available for New Users as Expected on the Jobs Page "); 
		   	  }
			  
			    log.info("End TEST-------------------------- StudentsNoRecommendation------------------------");     
	    
 log.info("END=PASSED");
 StudentsNoRecommendation.addResultForTestCase("1214", TEST_CASE_PASSED_STATUS, ""); 	    
 
		  } 
   
		  catch (Exception e) {
			  log.info("END=FAILED");
	   StudentsNoRecommendation.addResultForTestCase("1214", TEST_CASE_FAILED_STATUS,  e.getMessage()); 	    
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
