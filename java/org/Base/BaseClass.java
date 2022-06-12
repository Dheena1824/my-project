package org.Base;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.LinkedList;
import java.util.List;
import java.util.Set;
import java.util.concurrent.TimeUnit;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.Alert;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;

import io.github.bonigarcia.wdm.WebDriverManager;

public class BaseClass {

	public static WebDriver driver;
	public static CharSequence shift;
	
	public static WebDriver browserLaunch(String bl) {
		if(bl.equalsIgnoreCase("chrome")) {
			WebDriverManager.chromedriver().setup();
			 driver=new ChromeDriver();
		}
		else if(bl.equalsIgnoreCase("firefox")) {
			WebDriverManager.firefoxdriver().setup();
		    driver=new FirefoxDriver();
		}
		else if(bl.equalsIgnoreCase("edge")) {    
			WebDriverManager.edgedriver().setup();
	        driver=new EdgeDriver();
		}
	    
		return driver;
		}

	public static void Maximize() {
		driver.manage().window().maximize();
	}
     
	public static void dynamicWait(int a) {
        driver.manage().timeouts().implicitlyWait(a, TimeUnit.SECONDS);
	} 
	
    public static void getUrl(String a) {
		driver.get(a);

	}

    public static void click(WebElement a) {
		a.click();

	}

    public static void getCurrentUrl() {
		driver.getCurrentUrl();

	}

    public static void getTitle() {
		driver.getTitle();

	}
    
    public static void quit() {
		driver.quit();
       
	}

    public static void getWindow(int a) {
		Set<String> allTabs = driver.getWindowHandles();
	    List<String> li=new LinkedList<String>();
	    li.addAll(allTabs);
	    driver.switchTo().window(li.get(a));
	}
     
    public static void takeScreenShot() throws IOException {
    	TakesScreenshot ts=(TakesScreenshot)driver;
		File file = ts.getScreenshotAs(OutputType.FILE);
		long filename = System.currentTimeMillis();
		File f=new File("C:\\Users\\ELCOT\\Documents\\eclipse\\Maven_testing\\src\\test\\resources\\ScreenShots\\"+filename+".jpg");
		FileUtils.copyFile(file, f);

	}
    
    public static void takeScreenShotWeblement(WebElement a) throws IOException {
        File ss2 = a.getScreenshotAs(OutputType.FILE);		
		long filename = System.currentTimeMillis(); 
		File s1=new File("C:\\Users\\ELCOT\\Documents\\eclipse\\Maven_testing\\src\\test\\resources\\ScreenShots\\"+filename+".jpg");
		FileUtils.copyFile(ss2, s1);

	}
    
    public static void moveToElement(WebElement tar) {
		Actions a=new Actions(driver);
        a.moveToElement(tar).perform();
	}

    public static void dragAndDrop(WebElement src,WebElement tar) {
		Actions a=new Actions(driver);
        a.dragAndDrop(src, tar).perform();
	}
    
    public static void clickActions(WebElement tar) {
		Actions a=new Actions(driver);
        a.click(tar).perform();
    }
    public static void contextClick(WebElement e) {
       Actions a=new Actions(driver);
	   a.contextClick(e).perform();
    }

	public static void doubleClick(WebElement e) {
	    Actions a=new Actions(driver);
		a.doubleClick(e).perform();		
			
	}
	
	public static void keyDownShift(WebElement e) {
	    Actions a=new Actions(driver);
		a.keyDown(e,shift);
			
	}
	
	public static void keyDownShift() {
	    Actions a=new Actions(driver);
		a.keyDown(shift);
			
	}
	
	public static void keyUpShift(WebElement e) {
	    Actions a=new Actions(driver);
		a.keyUp(e,shift);
			
	}
	
	public static void keyUpShift() {
	    Actions a=new Actions(driver);
		a.keyUp(shift);
			
	}
	
    public static void selectByIndex(WebElement e,int a) {
	   Select s=new Select(e); 
	   s.selectByIndex(a);
   }
    public static void selectByVisibleText(WebElement e,String a) {
 	   Select s=new Select(e); 
 	   s.selectByVisibleText(a);
    }
    public static void selectByValue(WebElement e,String a) {
  	   Select s=new Select(e); 
  	   s.selectByValue(a);
  	   
     }
    public static void deSelectByIndex(WebElement e,int a) {
 	   Select s=new Select(e); 
 	   s.selectByIndex(a);
    }
     public static void deSelectByVisibleText(WebElement e,String a) {
  	   Select s=new Select(e); 
  	   s.selectByVisibleText(a);
     }
     public static void deSelectByValue(WebElement e,String a) {
   	   Select s=new Select(e); 
   	   s.selectByValue(a);
   	   
      }
     public static void down(int a) throws AWTException {
		Robot r=new Robot();
		for(int i=0;i<a;i++) {
        r.keyPress(KeyEvent.VK_DOWN);
        r.keyRelease(KeyEvent.VK_DOWN);
		}
	}
    
     public static void enter() throws AWTException {
    	 Robot r=new Robot();
    	 r.keyPress(KeyEvent.VK_ENTER);
         r.keyRelease(KeyEvent.VK_ENTER);

	}
     
     public static void pressShift() throws AWTException {
    	 Robot r=new Robot();
    	 r.keyPress(KeyEvent.VK_SHIFT);

	}
     
     public static void repleaseShift() throws AWTException {
    	 Robot r=new Robot();
    	 r.keyRelease(KeyEvent.VK_SHIFT);

	}
     
     public static void controlC() throws AWTException {
    	 Robot r=new Robot();
    	 r.keyPress(KeyEvent.VK_CONTROL);
    	 r.keyPress(KeyEvent.VK_C);
    	 r.keyRelease(KeyEvent.VK_C);
         r.keyRelease(KeyEvent.VK_CONTROL);

	}
     
     public static void controlS() throws AWTException {
    	 Robot r=new Robot();
    	 r.keyPress(KeyEvent.VK_CONTROL);
    	 r.keyPress(KeyEvent.VK_S);
    	 r.keyRelease(KeyEvent.VK_S);
         r.keyRelease(KeyEvent.VK_CONTROL);

	}
     
     public static void controlZ() throws AWTException {
    	 Robot r=new Robot();
    	 r.keyPress(KeyEvent.VK_CONTROL);
    	 r.keyPress(KeyEvent.VK_Z);
    	 r.keyRelease(KeyEvent.VK_Z);
         r.keyRelease(KeyEvent.VK_CONTROL);

	}
     
     public static void controlX() throws AWTException {
    	 Robot r=new Robot();
    	 r.keyPress(KeyEvent.VK_CONTROL);
    	 r.keyPress(KeyEvent.VK_X);
    	 r.keyRelease(KeyEvent.VK_X);
         r.keyRelease(KeyEvent.VK_CONTROL);

	}
     
     public static void controlV() throws AWTException {
    	 Robot r=new Robot();
    	 r.keyPress(KeyEvent.VK_CONTROL);
    	 r.keyPress(KeyEvent.VK_V);
    	 r.keyRelease(KeyEvent.VK_V);
         r.keyRelease(KeyEvent.VK_CONTROL);

	}
     
     public static void scrollDown(WebElement a) {
		JavascriptExecutor js=(JavascriptExecutor)driver;
        js.executeScript("arguments[0].scrollIntoView('true')", a);
     }

     public static void scrollUp(WebElement a) {
 		JavascriptExecutor js=(JavascriptExecutor)driver;
         js.executeScript("arguments[0].scrollIntoView('false')", a);
      }

     public static void setAttributeJava(String txt,WebElement a1) {
 		JavascriptExecutor js=(JavascriptExecutor)driver;
         js.executeScript("arguments[0].setAttribute('value','+txt+')", a1);
      }
     
     public static void getAttributeJava(WebElement a) {
  		JavascriptExecutor js=(JavascriptExecutor)driver;
          Object executeScript = js.executeScript("return arguments[0].getAttribute('value')", a);
          String value =(String)executeScript;
          System.out.println(value);
       }
     public static void clickJava(WebElement a) {
  		JavascriptExecutor js=(JavascriptExecutor)driver;
          js.executeScript("arguments[0].click()", a);
       }
     
     public static String getText(WebElement a) {
		String text = a.getText();
		return text;
        
	}
      
     public static String getAttribute(WebElement a,String e) {
		String attribute = a.getAttribute(e);
		return attribute;
        
	}
     
    public static void sleep() throws InterruptedException {
		Thread.sleep(3000);

	}
    public static void sts(String a) {
		System.out.println(a);
    }

	public static void sendKeys(WebElement a,Keys s,String e) {
		a.sendKeys(s,e);

	}
	
	public static void sendKeys(WebElement a,String e,Keys s) {
		a.sendKeys(e,s);

	}
    
	public static void sendKeys(WebElement a,String e) {
		a.sendKeys(e);

	}
     
	public static void acceptAlert() {
		Alert a = driver.switchTo().alert();
        a.accept();
		
	}
     
	public static void dismissAlert() {
		Alert a = driver.switchTo().alert();
        a.dismiss();
		
	}
     
	public static void sendkeysAlert(String e) {
		Alert a = driver.switchTo().alert();
        a.sendKeys(e);
		
	}
     
	public static String getTextAlert() {
		Alert a = driver.switchTo().alert();
        String text = a.getText();
		return text;
		
	}
     
	public static void frameIndex(int a) {
		driver.switchTo().frame(a);

	}
    
	public static void frameIdOrName(String a) {
		driver.switchTo().frame(a);

	}
	
	public static void frameWebelement(WebElement a) {
		driver.switchTo().frame(a);

	}
    
	public static void parentFrame() {
		driver.switchTo().parentFrame();

	}
     
	public static void defaultContent() {
		driver.switchTo().defaultContent();

	}
     
	public static void clear(WebElement a) {
		a.clear();
             
	}
     
    public static void back() {
		driver.navigate().back();

	}   
     
    public static void forward() {
		driver.navigate().forward();

	} 
     
    public static void refresh() {
		driver.navigate().refresh();

	}
     
    public static void toUrl(String a) {
		driver.navigate().to(a);

	}
    
    public static String readExcel(String fileName,String sheet,int row,int cel) throws IOException {
		File  f=new File("C:\\Users\\ELCOT\\Documents\\eclipse\\Maven_testing\\src\\test\\resources\\xcell\\"+fileName+".xlsx");
        FileInputStream st=new FileInputStream(f);
        Workbook w=new XSSFWorkbook(st);		
		Sheet s = w.getSheet(sheet);
		Row r = s.getRow(row);
		Cell cell = r.getCell(cel);
		int cellType = cell.getCellType();
		String value=null;
		if(cellType==1) {
			value = cell.getStringCellValue();
		}
		else {
			if(DateUtil.isCellDateFormatted(cell)) {
				Date dateCellValue = cell.getDateCellValue();
				SimpleDateFormat sd=new SimpleDateFormat("dd-MMM-yyyy");
				value = sd.format(dateCellValue);
			}
			else {
				double numericCellValue = cell.getNumericCellValue();
				long  l=(long)numericCellValue;
				value = String.valueOf(l);
			}
		}
		return value;
	}
    
    public static void createExcel(String fileName,String sheetName,int row,int cel,String data) throws IOException {
		File f=new File("C:\\Users\\ELCOT\\Documents\\eclipse\\Maven_testing\\src\\test\\resources\\xcell\\"+fileName+".xlsx");
        Workbook w=new XSSFWorkbook();
        Sheet s = w.createSheet(sheetName);
        Row r = s.createRow(row);
        Cell cell = r.createCell(cel);
        cell.setCellValue(data);
        
        FileOutputStream ot=new FileOutputStream(f);
        w.write(ot);
	} 
    
     public static void addDataExcel(String fileName,String sheetName,int row,int cel,String data) throws IOException {
    	 File f=new File("C:\\Users\\ELCOT\\Documents\\eclipse\\Maven_testing\\src\\test\\resources\\xcell\\"+fileName+".xlsx");
         FileInputStream st=new FileInputStream(f);
    	 Workbook w=new XSSFWorkbook(st);
         Sheet s = w.getSheet(sheetName);
         Row r = s.createRow(row);
         Cell cell = r.createCell(cel);
         cell.setCellValue(data);
         
         FileOutputStream ot=new FileOutputStream(f);
         w.write(ot);
         
	}
     
     public static void updateExcel(String fileName,String sheetName,int row,int cel,String toData) throws IOException {
    	 File f=new File("C:\\Users\\ELCOT\\Documents\\eclipse\\Maven_testing\\src\\test\\resources\\xcell\\"+fileName+".xlsx");
         FileInputStream st=new FileInputStream(f);
    	 Workbook w=new XSSFWorkbook(st);
         Sheet s = w.getSheet(sheetName);
         Row r = s.getRow(row);
         Cell cell = r.getCell(cel);
         cell.setCellValue(toData);
         FileOutputStream ot=new FileOutputStream(f);
         w.write(ot);
         
	}
     
     
     
     













}
