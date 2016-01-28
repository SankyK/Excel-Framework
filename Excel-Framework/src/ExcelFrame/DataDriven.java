package ExcelFrame;

import org.testng.annotations.Test;

import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.awt.image.BufferedImage;
import java.awt.image.DataBuffer;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

import javax.imageio.ImageIO;

import org.apache.commons.io.FileUtils;
//import org.apache.log4j.*;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.Assert;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.format.CellFormat;
import jxl.format.Colour;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableCell;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

public class DataDriven {
	Date date1 = new Date();
	SimpleDateFormat dateFormat1 = new SimpleDateFormat("yyyy_MM_dd_HH_mm_ss"); // Output excel sheet named with Date format 
	String Fname = dateFormat1.format(date1);	
    String fname = dateFormat1.format(date1);	
	// Create webdriver interface reference as fields of test class
	
	public WebDriver driver;
	public WebDriverWait wait;
	static int copy;
	String val= "Pass";
	String val1 = "Fail";
	
	@BeforeClass
	public void testSetup() {
		driver=new FirefoxDriver();
		driver.manage().window().maximize();
		wait = new WebDriverWait(driver, 5);
	}
	
	@AfterClass
	public void tearDown() {
		driver.quit();
	}
	

	// Call the test method
	@Test(dataProvider="data-provider")
	
	//Calling all test data from Input excel sheet
	public final void testClick(String exe,String URL,String Gmail,String Images,String Gapps,String SignIn,String Search,String Keys,String GooglesS) throws InterruptedException, ParseException, IOException 
	{
					
	Date d1 = new Date();
	SimpleDateFormat dateFormat1 = new SimpleDateFormat("dd-MM-yyyy");
	String Fd = dateFormat1.format(d1);
	
	try{
	    if(exe.equals("Y")||exe.equals("Yes")||exe.equals("y")||exe.equals("yes"))
	        {
	    	try{
	    	    driver.get(URL);
	    	    Thread.sleep(1000);
	    	    
	    	    driver.findElement(By.xpath(Gmail)).click();
	    	    Thread.sleep(1000);
	    	    
	    	    driver.get(URL);
	    	    Thread.sleep(1000);
	    	    	    	    
	    	    driver.findElement(By.xpath(Images)).click();
	    	    Thread.sleep(1000);
				
	    	    driver.get(URL);
	    	    Thread.sleep(1000);	    	
	    	    
	    	    driver.findElement(By.xpath(Gapps)).click();
	    	    Thread.sleep(1000);
				
	    	    driver.findElement(By.xpath(SignIn)).click();
	    	    Thread.sleep(1000);
	    	    
	    	    driver.get(URL);
	    	    Thread.sleep(1000);
	    	    
	    	    driver.findElement(By.xpath(Search)).sendKeys(Keys);
	    	    Thread.sleep(1000);
	    	    
	    	    driver.findElement(By.xpath(GooglesS));
	    	    Thread.sleep(3000);
					    
							writeexcel("Pass",Colour.GREEN);

	        }catch(Exception e){
	        	
	        }
	        }else{
				writeexcel("Fail",Colour.RED);
				
			}
	
	}	catch(Exception e)
	{
		
	}
	}

@DataProvider(name = "data-provider")
public String[][] data() throws Exception {
		String[][] arrayObject = getExcelData("C:\\Users\\Sanky\\Desktop\\Excel\\TestData\\Exceldata.xls","Sheet1");
		return arrayObject;
	}

	/**
	 * @param File Name
	 * @param Sheet Name
	 * @return
	 * @throws Exception 
	 * @throws BiffException 
	 */
	public String[][] getExcelData(String fileName, String sheetName) throws BiffException, Exception {
		String s2 = null;
		String[][] arrayExcelData = null;
		try {
		FileInputStream fs = new FileInputStream(fileName);
				Workbook wb = Workbook.getWorkbook(fs);
			Sheet sh = wb.getSheet(sheetName);

			int totalNoOfCols = sh.getRow(0).length;
			int totalNoOfRows = sh.getRows();
			
			arrayExcelData = new String[totalNoOfRows-1][totalNoOfCols];
			
			for (int i= 1 ; i < totalNoOfRows; i++) {

			for (int j=0; j < totalNoOfCols; j++) {
				arrayExcelData[i-1][j] = sh.getCell(j, i).getContents();
			}

			}
			} catch (FileNotFoundException e) {
			e.printStackTrace();
				}
		System.out.println("\n"+arrayExcelData);
		return arrayExcelData;
		
		
	}
	
	public void writeexcel(String s, Colour colour) throws Exception
	{
		FileInputStream fs = new FileInputStream("C:\\Users\\Sanky\\Desktop\\Excel\\TestData\\Exceldata.xls");
		Workbook  wb = Workbook.getWorkbook(fs);
		Fname= "C:\\Users\\Sanky\\Desktop\\Excel\\TestData\\" + Fname+ ".xls"; // Creating Output excel sheet
	    WritableWorkbook wkr = Workbook.createWorkbook(new File(Fname), wb);
	  Sheet sh = wb.getSheet(0);
		 WritableSheet getsht = wkr.getSheet(0);
		
	int totalNoOfCols = sh.getRow(0).length;
	
	int totalNoOfRows = sh.getRows();
	for (int i= 1 ; i < totalNoOfRows; i++) {
		
	    Label label = new Label(25, i,s, getCellFormat (colour));
			 getsht.addCell(label);
			}
	
	wkr.write();
	wkr.close();
	}
	
	private static CellFormat getCellFormat(Colour colour) throws Exception {
		WritableFont cellFont = new WritableFont(WritableFont.TIMES, 16);
	    WritableCellFormat cellFormat = new WritableCellFormat(cellFont);
	    cellFormat.setBackground(colour);
	    return cellFormat;
	
	}

}