package day3;

import java.util.List;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.By.ById;
import org.openqa.selenium.By.ByLinkText;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.Test;

import com.aventstack.extentreports.ExtentReporter;
import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.Status;
import com.aventstack.extentreports.markuputils.ExtentColor;
import com.aventstack.extentreports.markuputils.MarkupHelper;
import com.aventstack.extentreports.reporter.ExtentHtmlReporter;

public class ApachePoi_demo1 {
  @Test
  public void f() throws IOException, InterruptedException {
	  File src=new File("C:\\Users\\training_b6B.01.16\\Desktop\\test_data.xlsx");
	  FileInputStream fis= new FileInputStream(src);
	  XSSFWorkbook WB = new XSSFWorkbook(fis);
	  XSSFSheet SH=WB.getSheetAt(0);
	  System.out.println("Frist row number"+SH.getFirstRowNum());
	  System.out.println("last row number"+SH.getLastRowNum());
	  int rowCount=SH.getLastRowNum()-SH.getFirstRowNum();
	  System.out.println("the total rowcount is"+rowCount);
	  for(int i=1;i<=rowCount;i++) {
	  System.out.println(SH.getRow(i).getCell(0).getStringCellValue()+"\t\t\t"+SH.getRow(i).getCell(1).getStringCellValue());
	  
	  
	  System.setProperty("webdriver.chrome.driver","C:\\Users\\training_b6B.01.16\\Desktop\\BrowserDrivers\\chromedriver.exe");
	  ChromeDriver driver=new ChromeDriver();
	  String url="http://10.232.237.143:443/TestMeApp/fetchcat.htm";
	  driver.get(url);
	  driver.findElement(By.linkText("SignIn")).click();
	  Thread.sleep(5000);
	  driver.findElement(By.name("userName")).sendKeys(SH.getRow(i).getCell(0).getStringCellValue());
	  driver.findElement(By.id("password")).sendKeys(SH.getRow(i).getCell(1).getStringCellValue());
	  Thread.sleep(5000);
	  driver.findElement(By.name("Login")).click();
	  driver.findElement(By.xpath("//*[@id=\"menu3\"]/li[4]/a/span")).click();
	  WebElement objtable=driver.findElement(By.xpath("/html/body/b/section/div"));
	  List<WebElement> Allrows=objtable.findElements(By.tagName("tr"));
	  for(int n=1;n<Allrows.size();n++) {
		  List<WebElement> Cells=Allrows.get(n).findElements(By.tagName("td"));
		 System.out.println("Order: "+Cells.get(0).getText());
		 System.out.println("Order: "+Cells.get(2).getText());
		  }
		  
	  }
	  
/*writing into excel
 XSSFRow header=SH.getRow(0);
 XSSFCell header2=header.createCell(2);
 header2.setCellValue("status");
 SH.getRow(1).createCell(2).setCellValue("PASS");
 SH.getRow(2).createCell(2).setCellValue("fail");
 FileOutputStream fos=new FileOutputStream(src);
 WB.write(fos);*/
	  
	  /*ExtentHtmlReporter reporter=new ExtentHtmlReporter("C:\\\\Users\\\\training_b6B.01.16\\\\Desktop\\mynextextent reporter.html\\");
	  ExtentReports extent=new ExtentReports();
	  extent.attachReporter(reporter);
	  ExtentTest logger=extent.createTest("DemoWebShop");
	  logger.log(Status.INFO,"Apache poi is used in this script");
	  logger.log(Status.PASS,"excel date is reading is done successfully");
	  logger.log(Status.FAIL, MarkupHelper.createLabel("THIS TEST CASE IS FAIL", ExtentColor.BLUE));
	  extent.flush();
	  driver.close();*/
	 
	  
	  }
  }
 
 


