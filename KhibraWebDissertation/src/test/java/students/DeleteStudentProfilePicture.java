 package students;

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

import utils.LoginUtility;
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
 
public class DeleteStudentProfilePicture extends LoginUtility{
	 
		
	@BeforeMethod
    public void setUp1() throws Exception {
	  log.info("START=DeleteStudentProfilePicture--------------------------------------------------------------");

	  Loginutilityclass();
    
    	
    		
}
	
	
	  @SuppressWarnings("rawtypes")
	@Test
	
	public void testDeleteStudentProfilePicture() throws Exception {
		  try{
			   
			  
			   
			  Thread.sleep(2000);
			  String filePath = System.getProperty("user.dir");
		 	  FileInputStream fis = new FileInputStream(filePath+"/config/Testdata.xlsx");
	   
			    XSSFWorkbook workbook = new XSSFWorkbook(fis); 
				  
				  XSSFSheet sheets = workbook.getSheet("Students");
				 
				  
				  Row studentphoto = sheets.getRow(1);
				  Cell cell1 = studentphoto.getCell(1);
				  String Studentphoto  = cell1.getStringCellValue();
 
			  driver.findElement(By.xpath(objectrepo.getProperty("MenuDropdown"))).click();	 
			  log.info("Menu Dropdown Clicked");
			  driver.findElement(By.xpath(objectrepo.getProperty("MyProfileMenu"))).click();	 
			  log.info("My profile Selected");
			  driver.findElement(By.xpath(objectrepo.getProperty("EditBasicInfo"))).click();	 
Thread.sleep(2000);
log.info("Edit Basic Info Clicked");
			  
			  
			  Thread.sleep(2000);
driver.findElement(By.xpath(objectrepo.getProperty("UploadNewPhotoButton"))).sendKeys(System.getProperty("user.dir")+Studentphoto);
log.info("Image Selected ");
Thread.sleep(2000);


driver.findElement(By.xpath(objectrepo.getProperty("DoneButton"))).click();	 
log.info("DONE BUTTON CLICKED ");
 
driver.findElement(By.xpath(objectrepo.getProperty("SavedButtons"))).click();	
log.info("Save BUTTON CLICKED ");
driver.findElement(By.xpath(objectrepo.getProperty("EditBasicInfo"))).click();	 
Thread.sleep(2000);
log.info("Edit Basic Info Clicked");

			  Thread.sleep(2000);
	 
driver.findElement(By.xpath(objectrepo.getProperty("DeletePhotoButton"))).click();	 
log.info("Delete BUTTON CLICKED ");
Thread.sleep(2000);
driver.findElement(By.xpath(objectrepo.getProperty("SavedButtons"))).click();	
log.info("Save BUTTON CLICKED and Photo is deleted ");


Thread.sleep(2000);
 

Logs =   " log.info(\"Menu Dropdown  Clicked\");\r\n"
		+ "		  log.info(\"My profile Selected\");\r\n"
		+ "		  log.info(\"Edit Basic Info Clicked\");\");\r\n"
		+ "		  log.info(\"Delete BUTTON CLICKED \");\r\n"
		+ "		  log.info(\"Save BUTTON CLICKED and Photo is deleted\");\r\n"
		 
			; 


			    log.info("End TEST-------------------------- DeleteStudentProfilePicture------------------------");     
	    
			    log.info("END=PASSED");
			    DeleteStudentProfilePicture.addResultForTestCase("1192", TEST_CASE_PASSED_STATUS, ""); 	    		     
		  } 

		  catch (Exception e) {
			  log.info("END=FAILED");
			  DeleteStudentProfilePicture.addResultForTestCase("1192", TEST_CASE_FAILED_STATUS,  e.getMessage()); 	    
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
