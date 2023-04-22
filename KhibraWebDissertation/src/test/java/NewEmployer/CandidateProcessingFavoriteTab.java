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
 
public class CandidateProcessingFavoriteTab extends EmployerUtility{
	 
		
	@BeforeMethod
    public void setUp1() throws Exception {
	  log.info("START=CandidateProcessingFavoriteTab--------------------------------------------------------------");

	  EmployerUtilityclass();
    
    	  
    		  		  
    		
}
	
	
	  @SuppressWarnings("rawtypes")
  @Test  
	
	public void testCandidateProcessingFavoriteTab() throws Exception {
		  try {
    		  Thread.sleep(2000);
    		  			  
    		  			  String filePath = System.getProperty("user.dir");
    		  		 	  FileInputStream fis = new FileInputStream(filePath+"/config/Testdata.xlsx");
    		  			    XSSFWorkbook workbook = new XSSFWorkbook(fis); 			  
    		  				  XSSFSheet sheets = workbook.getSheet("Employer");
    		  	 	  
    		  				  Row title = sheets.getRow(12);
    		  				  Cell cell1 = title.getCell(1);
    		  				  String JobTitle  = cell1.getStringCellValue();
    		  				  
    		  				 Row jobname = sheets.getRow(39);
			  				  Cell cell2 = jobname.getCell(1);
			  				  String Jobname  = cell2.getStringCellValue();
    		  				  
    		  				  
    		  				 
    		  				  driver.findElement(By.xpath(objectrepo.getProperty("EmpMenuDropdown"))).click();	 
    		  			  	  log.info("Dropdown  Clicked");
    		  			  	  Thread.sleep(2000);  
    		  			  
    		  driver.findElement(By.xpath(objectrepo.getProperty("EmpMyJobsDropdownMenu"))).click();	 
    		  	  log.info("Jobs Menu  Clicked");
    		  	  Thread.sleep(2000);
    		   
    			  
    		  	driver.findElement(By.xpath(objectrepo.getProperty("OpenJobsMenu"))).click();	 
    		  	log.info("Open Menu  Clicked");
    		  	Thread.sleep(2000);
    		  	
    			  
		  		 driver.findElement(By.xpath(objectrepo.getProperty("JobSearchBox"))).sendKeys(Jobname);
		  		  log.info("Job Title entered into Search Box");
		  		  Thread.sleep(2000);
		  		  
		  		  driver.findElement(By.xpath(objectrepo.getProperty("SearchButton"))).click();	 
		  		  log.info("Search Button  Clicked");
		  		  Thread.sleep(2000);
		  		  
		  		  driver.getPageSource().contains(JobTitle);
		  		 log.info("Job Searched was found");
    		  	  
    		  	  driver.findElement(By.xpath(objectrepo.getProperty("FirstJobRow"))).click();	 
    		  	  log.info("Jobs Selected");
    		  	  Thread.sleep(6000);
    		  	driver.manage().timeouts().implicitlyWait(60, TimeUnit.SECONDS);
    		  /*	  
    		  	 driver.findElement(By.xpath(objectrepo.getProperty("RecommendedTAB"))).click();	 
   		  	  log.info("Recommended TAB Selected");
   		  	  Thread.sleep(5000);
   		  	driver.manage().timeouts().implicitlyWait(60, TimeUnit.SECONDS);
   		  	  
   		   driver.findElement(By.xpath(objectrepo.getProperty("Checkbox"))).click();	 
	  		  log.info("CheckBox  Clicked to select a Talent ");
	  		  Thread.sleep(2000);
	  		  
	  		   driver.findElement(By.xpath(objectrepo.getProperty("ActionsButton"))).click();	 
		  		  log.info("Actions Buttons  Clicked");
		  		  Thread.sleep(2000);
	  		
  		  
		  		   driver.findElement(By.xpath(objectrepo.getProperty("FavoriteCandidate"))).click();	 
			  		  log.info("Favorite Button   Clicked ");
			  		 Thread.sleep(2000);
			  		   driver.findElement(By.xpath(objectrepo.getProperty("FavoriteSuccess"))).click();	 
			  		 log.info("Favorite success message is displayed  ");
			  		 Thread.sleep(2000);
			  		 */
		   		   Thread.sleep(2000);
   driver.findElement(By.xpath(objectrepo.getProperty("FavoriteTAB"))).click();	 
 				  		  log.info("Favorite TAB   Clicked ");
				  		  Thread.sleep(4000);	
				  	 
				  		  
			    driver.findElement(By.xpath(objectrepo.getProperty("Checkbox"))).click();	 
				 	  log.info("CheckBox  Clicked to select a Talent ");
				  Thread.sleep(2000);			  		  
				  		  
				  
		  		   driver.findElement(By.xpath(objectrepo.getProperty("ActionsButton"))).click();	 
			  		  log.info("Actions Buttons  Clicked");
			  		  Thread.sleep(2000);
			  		 
			  		  driver.findElement(By.xpath(objectrepo.getProperty("ProposeCandidate")));	
			  		 driver.findElement(By.xpath(objectrepo.getProperty("InterviewAction")));	
			  		 driver.findElement(By.xpath(objectrepo.getProperty("HireAction")));	
			  		 driver.findElement(By.xpath(objectrepo.getProperty("SendOfferAction")));	
			  		 driver.findElement(By.xpath(objectrepo.getProperty("UnfavoriteAction")));			  		 
			  		
			  		 log.info("All Actions Options Available on the Favorite TAB");
			  		   		  	  
			  		    
			  		  Logs =   " log.info(\"Employer Dropdown  Menu Clicked\");\r\n"
			  				+ "		  log.info(\"Jobs Menu  Clicked\");\r\n"
			  				+ "		  log.info(\"log.info(\"Job Title entered into Search Box and search Button Clickd\");\");\r\n"
			  				+ "		  log.info(\"Job Searched was found and verified\");\r\n"
			  				+ "		  log.info(\"Jobs Selected\");\r\n"
			  				+ "		  log.info(\"Favorite Button   Clicked\");\r\n"
			  				+ "		  log.info(\"Favorite success message is displayed \");\r\n"
			  				+ "		  log.info(\"Favorite TAB   Clicked \");\r\n"
			  				+ "		  log.info(\"CheckBox  Clicked to select a Talent \");\r\n"
			  				+ "		  log.info(\"Actions Buttons  Clicked\");\r\n"
			  				+ "		  log.info(\"All Actions Options Available on the Favorite TAB(Propose, Interview, Hire, SendOffer Unfavorite\");\r\n"
 
			  				; 
				  
					  		  
					  		  
			  		 
			  		 
			  		 
			  		 
			  		 
			  		 
			  		 
			  		 
			  		 
			    log.info("End TEST-------------------------- CandidateProcessingFavoriteTab------------------------");     
	    
			    log.info("END=PASSED");
			    CandidateProcessingAppliedTAB.addResultForTestCase("1159", TEST_CASE_PASSED_STATUS, ""); 	    
							   

								  } 
							  

		  catch (Exception e) {
			  log.info("END=FAILED");
	  CandidateProcessingAppliedTAB.addResultForTestCase("1159", TEST_CASE_FAILED_STATUS,  e.getMessage()); 	    
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
