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
import java.util.concurrent.TimeUnit;
import java.util.regex.Pattern;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
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
import org.openqa.selenium.interactions.Actions;
import org.testng.Assert;
import org.testng.AssertJUnit;
import org.testng.ITestResult;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import com.gargoylesoftware.htmlunit.ElementNotFoundException;

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
 
public class ChangeEmployerName extends EmployerUtility{
	 
		
	@BeforeMethod
    public void setUp1() throws Exception {
	  log.info("START=ChangeEmployerName--------------------------------------------------------------");

	  EmployerUtilityclass();
    
    	
    		
}
	
	
	  @SuppressWarnings("rawtypes")
	@Test
	
	public void testChangeEmployerName() throws Exception {
		  try {
			  
			  String filePath = System.getProperty("user.dir");
		 	  FileInputStream fis = new FileInputStream(filePath+"/config/Testdata.xlsx");
	   
			    XSSFWorkbook workbook = new XSSFWorkbook(fis); 
				  
				  XSSFSheet sheets = workbook.getSheet("Employer");
				  
				   
				  
				  Row empname = sheets.getRow(32);
				  Cell cell1 = empname.getCell(1);
				  String Empname  = cell1.getStringCellValue();
				  
		  
			  driver.findElement(By.xpath(objectrepo.getProperty("EmployerProfileMenu"))).click();	 
			  log.info("Employer Profile MenuClicked");
			  Thread.sleep(2000);
			  
	  driver.findElement(By.xpath(objectrepo.getProperty("EditEmployerIcon"))).click();	 
	 log.info("Edit Employer Icon Clicked");
	  
	 
	  driver.findElement(By.xpath(objectrepo.getProperty("UniNameTextbox"))).clear();	 

	  driver.findElement(By.xpath(objectrepo.getProperty("UniNameTextbox"))).sendKeys(Empname);	 
	 log.info("Employer Name Entered");
	 
	 
	  driver.findElement(By.xpath(objectrepo.getProperty("CompanyIndustryDropdown"))).click();	 
	 log.info("Company Industry  Clicked");
       
 	  Actions keyDown = new Actions(driver); 
	  keyDown.sendKeys(Keys.chord(Keys.DOWN, Keys.ENTER)).perform();
			   
	  Thread.sleep(2000);  
	 
	 log.info("Company Industry Option Selected");
		 Thread.sleep(2000); 
        
			   
		     driver.findElement(By.xpath(objectrepo.getProperty("SaveButton"))).click();
			  log.info("Save Button  Clicked ");
		    
			  driver.findElement(By.xpath(objectrepo.getProperty("UniProfileUpdateMessage")));	 
			  log.info("Employer Name and Industry Saved successfully ");
			
			 
			  
			  Logs =   " log.info(\"Employer Profile Clicked\");\r\n"
		  				+ "		  log.info(\"Edit Employer Icon Clicked\");\r\n"
		  				+ "		  log.info(\"Employer Name Entered\");\");\r\n"
		  				+ "		  log.info(\"Company Industry  Clicked\");\r\n"
		  				+ "		  log.info(\"Company Industry Option Selected\");\r\n"
		  				+ "		  log.info(\"Save Button  Clicked\");\r\n"
		  				+ "		  log.info(\"Employer Name and Industry Saved successfully\");\r\n"	  				 
	 	  				; 
			  
			  
			  
			  
			  
			    log.info("End TEST--------------------------ChangeEmployerName------------------------");     
	    
			    log.info("END=PASSED");
			 	  
			    ChangeEmployerName.addResultForTestCase("1168", TEST_CASE_PASSED_STATUS, ""); 	    
		 		   
			    
		  } 

		  catch (Exception e) {
			  log.info("END=FAILED");
   ChangeEmployerName.addResultForTestCase("1168", TEST_CASE_FAILED_STATUS,  e.getMessage()); 	    
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
