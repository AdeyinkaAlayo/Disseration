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
 
public class EmpFilterTalents extends EmployerUtility{
	 
		
	@BeforeMethod
    public void setUp1() throws Exception {
	  log.info("START=EmpFilterTalents--------------------------------------------------------------");

	  EmployerUtilityclass();
    
    	
    		
}
	
	
	  @SuppressWarnings("rawtypes")
	@Test
	
	public void testEmpFilterTalents() throws Exception {
		  
		  try {
    		  Thread.sleep(2000);
    		  
    		  String filePath = System.getProperty("user.dir");
		 	  FileInputStream fis = new FileInputStream(filePath+"/config/Testdata.xlsx");
			    XSSFWorkbook workbook = new XSSFWorkbook(fis); 			  
				  XSSFSheet sheets = workbook.getSheet("Employer");
	 	  
				  Row name = sheets.getRow(13);
				  Cell cell1 = name.getCell(1);
				  String Name  = cell1.getStringCellValue();
				  
				  Row wrongname = sheets.getRow(14);
				  Cell cell2 = wrongname.getCell(1);
				  String Wrongname  = cell2.getStringCellValue();
				  
	     
    		  driver.findElement(By.xpath(objectrepo.getProperty("EmpTalentMenu"))).click();	 
		  	  log.info("Talent Menu  Clicked");
		  	  Thread.sleep(2000);
		  	  
	 
		  	  
		  	driver.findElement(By.xpath(objectrepo.getProperty("ShowFilterIcon"))).click();	 	  
		  	log.info("Show Filter Button  Clicked");
		  	  
		  	 
		  	 Actions keyDown = new Actions(driver);
		  	 
		  	 
		  	 driver.findElement(By.xpath(objectrepo.getProperty("SkillsFilter"))).click();	 
		  	 Thread.sleep(2000);
		  	 
		  	 keyDown.sendKeys(Keys.chord(Keys.DOWN, Keys.ENTER)).perform();
		  	 log.info("Skills Filter Seleted ");
		  	 
		  	driver.findElement(By.xpath(objectrepo.getProperty("UniFilter"))).click();	 
		  	 Thread.sleep(2000);
		  	 
 		  	   
		  	 keyDown.sendKeys(Keys.chord(Keys.DOWN,Keys.DOWN,Keys.DOWN, Keys.ENTER)).perform();
		  	 log.info("University Filter Seleted ");

		  	  
		  	 Thread.sleep(2000);
		  	 driver.findElement(By.xpath(objectrepo.getProperty("ApplyFilterButton"))).click();

		  	 Thread.sleep(4000);
 
		  	 
		  	 log.info("Apply Button Clicked and   values displayed");
		  	 
		 	driver.findElement(By.xpath(objectrepo.getProperty("ShowFilterIcon"))).click();	 	  
		  	log.info("Show Filter Button  Clicked");
		  	 
		  	 driver.findElement(By.xpath(objectrepo.getProperty("ClearFilters"))).click();
		  	 Thread.sleep(2000);
		  	 log.info("Clear Filter Button Clicked");
		  	
		  	
		  	
		  	
		  	
		  	
		  	
		  	
		  	
		  	
		  	
		  	
		  	
		  	  
 driver.findElement(By.xpath(objectrepo.getProperty("ExperienceFilter"))).click();	 
 Thread.sleep(2000);
 
 //driver.findElement(By.xpath(objectrepo.getProperty("ExperienceBox"))).click();

 keyDown.sendKeys(Keys.chord(Keys.DOWN,Keys.DOWN,Keys.DOWN, Keys.ENTER)).perform();
 log.info("Experience Filter Seleted ");

 driver.findElement(By.xpath(objectrepo.getProperty("ApplyFilterButton"))).click();
 
 Thread.sleep(2000);
 
	driver.findElement(By.xpath(objectrepo.getProperty("ShowFilterIcon"))).click();	 	  
  	log.info("Show Filter Button  Clicked");
  	
  	
 driver.findElement(By.xpath(objectrepo.getProperty("BadgeFilterCount")));
 
 log.info("Apply Filter Button Clicked and Experience values displayed");
 
 
 driver.findElement(By.xpath(objectrepo.getProperty("ClearFilters"))).click();
 Thread.sleep(2000);

 
 
 
 driver.findElement(By.xpath(objectrepo.getProperty("SkillsFilter"))).click();	 
 Thread.sleep(2000);
 
 keyDown.sendKeys(Keys.chord(Keys.DOWN,Keys.DOWN, Keys.ENTER)).perform();
 log.info("Skills Filter Seleted ");

 driver.findElement(By.xpath(objectrepo.getProperty("ApplyFilterButton"))).click();

 Thread.sleep(2000);
 
	driver.findElement(By.xpath(objectrepo.getProperty("ShowFilterIcon"))).click();	 	  
  	log.info("Show Filter Button  Clicked");
  	
  	
 driver.findElement(By.xpath(objectrepo.getProperty("BadgeFilterCount")));
 
 log.info("Apply Filter Button Clicked and Skills values displayed");
 driver.findElement(By.xpath(objectrepo.getProperty("ClearFilters"))).click();
 Thread.sleep(2000);
  
 log.info("Clear Filter Button Clicked");
 
 
 
 driver.findElement(By.xpath(objectrepo.getProperty("UniFilter"))).click();	 
 Thread.sleep(2000);
 
 //driver.findElement(By.xpath(objectrepo.getProperty("UniBox"))).click();
   
 keyDown.sendKeys(Keys.chord(Keys.DOWN,Keys.DOWN,Keys.DOWN, Keys.ENTER)).perform();
 log.info("University Filter Seleted ");

 driver.findElement(By.xpath(objectrepo.getProperty("ApplyFilterButton"))).click();

 Thread.sleep(2000);
 
	driver.findElement(By.xpath(objectrepo.getProperty("ShowFilterIcon"))).click();	 	  
  	log.info("Show Filter Button  Clicked");
  	
  	
 driver.findElement(By.xpath(objectrepo.getProperty("BadgeFilterCount")));
 
 log.info("Apply Filter Button Clicked and University values displayed");
 
 driver.findElement(By.xpath(objectrepo.getProperty("ClearFilters"))).click();
 Thread.sleep(2000);
 log.info("Clear Filter Button Clicked");
 

 
 driver.findElement(By.xpath(objectrepo.getProperty("MajorFilter"))).click();	 
 Thread.sleep(2000);
 
// driver.findElement(By.xpath(objectrepo.getProperty("MajorBox"))).click();
   
 keyDown.sendKeys(Keys.chord(Keys.DOWN,Keys.DOWN,Keys.DOWN, Keys.ENTER)).perform();
 log.info("Major Filter Seleted ");

 driver.findElement(By.xpath(objectrepo.getProperty("ApplyFilterButton"))).click();

 Thread.sleep(4000);
 
	driver.findElement(By.xpath(objectrepo.getProperty("ShowFilterIcon"))).click();	 	  
  	log.info("Show Filter Button  Clicked");
  	
  	
 driver.findElement(By.xpath(objectrepo.getProperty("BadgeFilterCount")));
 
 log.info("Apply Filter Button Clicked and Major values displayed");
 
 driver.findElement(By.xpath(objectrepo.getProperty("ClearFilters"))).click();
 Thread.sleep(2000);
 log.info("Clear Filter Button Clicked");
 
 
 
 driver.findElement(By.xpath(objectrepo.getProperty("DegreeFilter"))).click();	 
 Thread.sleep(2000);
 
// driver.findElement(By.xpath(objectrepo.getProperty("DegreeBox"))).click();
   
 keyDown.sendKeys(Keys.chord(Keys.DOWN,Keys.DOWN,Keys.DOWN, Keys.ENTER)).perform();
 log.info("Degree Filter Seleted ");

 driver.findElement(By.xpath(objectrepo.getProperty("ApplyFilterButton"))).click();

 Thread.sleep(4000);
 
	driver.findElement(By.xpath(objectrepo.getProperty("ShowFilterIcon"))).click();	 	  
  	log.info("Show Filter Button  Clicked");
  	
  	
 driver.findElement(By.xpath(objectrepo.getProperty("BadgeFilterCount")));
 
 log.info("Apply Filter Button Clicked and Degree values displayed");
 
 driver.findElement(By.xpath(objectrepo.getProperty("ClearFilters"))).click();
 Thread.sleep(2000);
 log.info("Clear Filter Button Clicked");
 
 
 driver.findElement(By.xpath(objectrepo.getProperty("GPAFilter"))).click();	 
 Thread.sleep(2000);
 
 // driver.findElement(By.xpath(objectrepo.getProperty("GPABox"))).click();
   
 keyDown.sendKeys(Keys.chord(Keys.UP,Keys.DOWN,Keys.UP, Keys.ENTER)).perform();
 log.info("Degree Filter Seleted ");
 Thread.sleep(2000);
 driver.findElement(By.xpath(objectrepo.getProperty("ApplyFilterButton"))).click();

 Thread.sleep(4000);
 
	driver.findElement(By.xpath(objectrepo.getProperty("ShowFilterIcon"))).click();	 	  
  	log.info("Show Filter Button  Clicked");
  	
  	
 driver.findElement(By.xpath(objectrepo.getProperty("BadgeFilterCount")));
 
 log.info("Apply   Button Clicked and GPA values displayed");
 
 driver.findElement(By.xpath(objectrepo.getProperty("ClearFilters"))).click();
 Thread.sleep(2000);
 log.info("Clear Filter Button Clicked");
 
 
 
 driver.findElement(By.xpath(objectrepo.getProperty("GradYearFilter"))).click();	 
 Thread.sleep(2000);
 log.info("Grad Year Filter clicked ");
 // driver.findElement(By.xpath(objectrepo.getProperty("GradYearBox"))).click();
   
 keyDown.sendKeys(Keys.chord(Keys.DOWN,Keys.DOWN,Keys.DOWN, Keys.ENTER)).perform();
 log.info("Grad Year Filter Seleted ");
 Thread.sleep(2000);
 driver.findElement(By.xpath(objectrepo.getProperty("ApplyFilterButton"))).click();

 Thread.sleep(2000);
 
	driver.findElement(By.xpath(objectrepo.getProperty("ShowFilterIcon"))).click();	 	  
  	log.info("Show Filter Button  Clicked");
  	
  	
 
 driver.findElement(By.xpath(objectrepo.getProperty("BadgeFilterCount")));
 
 log.info("Apply   Button Clicked and GPA values displayed");
 
 driver.findElement(By.xpath(objectrepo.getProperty("ClearFilters"))).click();
 Thread.sleep(2000);
 log.info("Clear Filter Button Clicked");
 

 
 
 
 
 
 
 
 
 
 
  
 
 
 driver.findElement(By.xpath(objectrepo.getProperty("NationalityFilter"))).click();	 
 Thread.sleep(2000);
 
 //driver.findElement(By.xpath(objectrepo.getProperty("CountryBox"))).click();
   
 keyDown.sendKeys(Keys.chord(Keys.DOWN,Keys.DOWN,Keys.DOWN, Keys.ENTER)).perform();
 log.info("Country Filter Seleted ");
 Thread.sleep(2000);
 
 driver.findElement(By.xpath(objectrepo.getProperty("ApplyFilterButton"))).click();

 Thread.sleep(2000);
  

 driver.findElement(By.xpath(objectrepo.getProperty("BadgeFilterCount")));
 
 log.info("Apply Button Clicked and Nationality  values displayed");
 
 driver.findElement(By.xpath(objectrepo.getProperty("ShowFilterIcon"))).click();	 	  
	log.info("Show Filter Button  Clicked");
 
 driver.findElement(By.xpath(objectrepo.getProperty("ClearFilters"))).click();
 Thread.sleep(2000);
 log.info("Clear Filter Button Clicked");
 
 
 
 
  
 
 driver.findElement(By.xpath(objectrepo.getProperty("GenderFilter"))).click();	 
 Thread.sleep(2000);
// driver.findElement(By.xpath(objectrepo.getProperty("GenderBox"))).click();	 
 Thread.sleep(2000);
 keyDown.sendKeys(Keys.chord(Keys.DOWN,Keys.DOWN,Keys.DOWN, Keys.ENTER)).perform();
 log.info("Gender Filter Seleted ");

 driver.findElement(By.xpath(objectrepo.getProperty("ApplyFilterButton"))).click();

 Thread.sleep(2000);
 
 driver.findElement(By.xpath(objectrepo.getProperty("ShowFilterIcon"))).click();	 	  
	log.info("Show Filter Button  Clicked");
 
 driver.findElement(By.xpath(objectrepo.getProperty("BadgeFilterCount")));
 
 log.info("Apply Filter Button Clicked and Gender values displayed");
 driver.findElement(By.xpath(objectrepo.getProperty("ClearFilters"))).click();
 Thread.sleep(2000);
  
 log.info("Clear Filter Button Clicked");
 
 
 


		   
			    log.info("End TEST-------------------------- EmpFilterTalents------------------------");     
	    
			    log.info("END=PASSED");
			     
		  EmpFilterTalents.addResultForTestCase("1174", TEST_CASE_PASSED_STATUS, ""); 	    
			     
		  } 

		  catch (Exception e) {
			  log.info("END=FAILED");
			  EmpFilterTalents.addResultForTestCase("1174", TEST_CASE_FAILED_STATUS,  e.getMessage()); 	    
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
