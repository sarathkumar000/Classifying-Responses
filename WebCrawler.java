package online;

import java.util.concurrent.TimeUnit;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;



public class Blurit {

	public static void main(String[] args) throws InterruptedException, IOException {
		// TODO Auto-generated method stub
		System.setProperty("webdriver.chrome.driver", "D:\\chromedriver_win32\\chromedriver.exe");
		WebDriver driver =new ChromeDriver();
		
		driver.get("https://www.blurtit.com/");
		driver.manage().timeouts().pageLoadTimeout(10, TimeUnit.SECONDS);
		driver.manage().timeouts().implicitlyWait(5,TimeUnit.SECONDS);
		driver.manage().window().maximize();
		
		// Things Related to Excel Sheet
		
		 //Create blank workbook
	      XSSFWorkbook workbook = new XSSFWorkbook();
	      
	      //Create a blank sheet
	      XSSFSheet spreadsheet = workbook.createSheet( "Forum Questions and answers");

	      //Create row object
	      XSSFRow row;
		try
		{
		// For loading all the Lists
		for(int i=0;i<250;i++)
		{
			System.out.println(i+"times");
		JavascriptExecutor js = (JavascriptExecutor) driver;
		js.executeScript("window.scrollTo(0, document.body.scrollHeight)");
		try
		{
		driver.findElement(By.xpath("//div[@class='long-button']")).click();
		}
		catch(Exception e)
		{
			break;
		}
		
		Thread.sleep(3000);
		}
		//After Loading all the lists we scrolling upto the page
		JavascriptExecutor js = (JavascriptExecutor) driver;
		js.executeScript("window.scrollTo(0,0)");
		
		//finding the No of list elemnts present
		List<WebElement>li=driver.findElements(By.xpath("//article"));
		int rowid = 0;
		for(int i=4000;i<li.size();i++)
		{
			
			
			// From here Taking qn and storing it in test by using xpath s3
			//By using x3 itself it is clicking that question so it will open another tab 
			String s="((//article)["+Integer.toString(i)+"]";
			String s2=s+"//div)[7]//li/a";
			String s3=s+"//div)[5]/a";
			String text=driver.findElement(By.xpath(s3)).getText();
			String selectLinkOpeninNewTab = Keys.chord(Keys.CONTROL,Keys.RETURN);
			try
			{
			driver.findElement(By.xpath(s3)).sendKeys(selectLinkOpeninNewTab);
			}
			catch(Exception e)
			{
				System.out.println("pass");
			}
			
			// After opening tab we are switching to that tab
			ArrayList<String> tabs = new ArrayList<String> (driver.getWindowHandles());
			try
			{
		    driver.switchTo().window(tabs.get(1));
			}
			catch(Exception e)
			{
				System.out.println("Tab Error");
			}
		    Thread.sleep(17000);
		    
		    // Now finding the number of answers present in that loaded page
		    List<WebElement>le=driver.findElements(By.xpath("//div[@itemprop='text']"));
		    System.out.println(text);
		    
		    // Now iterating the each answer and Printing
		    for(int j=1;j<le.size();j++)
		    {
		    	row = spreadsheet.createRow(rowid++);
		    	int cellid = 0;
		    	String k="";
		    	try
		    	{
		    	 k="(//div[@itemprop='text'])["+Integer.toString(j)+"]";
		    	}
		    	catch(Exception e)
		    	{
		    		System.out.println("Just GO Ahead");
		    	}
		    	if(j==1)
		    	{
		    		Cell cell = row.createCell(cellid);
		    		cell.setCellValue(text);
		    		
		    	}
		    	
		    		++cellid;
		    		Cell cell = row.createCell(cellid);	
		    		cell.setCellValue(driver.findElement(By.xpath(k)).getText());
		    	
		    	
		    	 
		    	System.out.println(driver.findElement(By.xpath(k)).getText());
		    	
		    }
		    
		  //div[@itemprop="text"]
		    driver.close();
		    
		    driver.switchTo().window(tabs.get(0));
		   // driver.findElement(By.xpath(s2)).sendKeys(Keys.CONTROL +"t");
		    System.out.println(i+"times executed");
			//String text=driver.findElement(By.xpath(s2)).getText();
			//System.out.println(text);
			//driver.findElement(By.xpath(s2)).click();
			Thread.sleep(5000);
			//driver.navigate().back();
		}
		
	}
	catch(Exception e)
	{
		System.out.println("A big Error occured So saving file to Data file");
	}
   finally
   {
		
		FileOutputStream out = new FileOutputStream(
		         new File("D:\\Data.xlsx"));
		      
		      workbook.write(out);
		      out.close();
		      System.out.println("Writesheet.xlsx written successfully");
		driver.quit();
   }
		
		
	}

}
