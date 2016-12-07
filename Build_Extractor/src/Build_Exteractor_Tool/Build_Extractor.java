package Build_Exteractor_Tool;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import java.util.concurrent.TimeUnit;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
//import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.remote.RemoteWebDriver;

public class Build_Extractor {

	static RemoteWebDriver driver;
	@Test
	public static void main(String[] args) throws Exception {
		Build_Extractor obj = new Build_Extractor();
		Date date = new Date();
        String dat = date.toString();
        String QA = "sprint";
        String stg = "Staging";
        String prd = "production";
        obj.writeExcel("output.xls", "output", dat);
        obj.writeExcel("output.xls", "output", QA);
		obj.BuildExtractor("http://myxfn.qa.xfinity.com/_status");
		obj.writeExcel("output.xls", "output", stg);
		obj.BuildExtractor("http://myxfn.staging.xfinity.com/_status");
		obj.writeExcel("output.xls", "output", prd);
		obj.BuildExtractor("http://my.xfinity.com/_status");
		
	}

	@Test
	public void BuildExtractor (String url) throws Exception{
		Build_Extractor obj1 = new Build_Extractor();
		System.setProperty("webdriver.chrome.driver", "C:\\Users\\266963\\workspace1\\Build_Extractor\\Drivers\\chromedriver.exe");
		driver =  new ChromeDriver();
		driver.get(url);
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
		WebElement build = driver.findElementByXPath("html/body/dl[1]/dd[4]");
		String name = build.getText();
		obj1.writeExcel("output.xls", "output", name);
		//System.out.println(name);
		driver.close();
		
		
	}
	
	@Test
	 public void writeExcel(String fileName,String sheetName,String dataToWrite) throws IOException{
		 
	        //Create a object of File class to open xlsx file
	 
	        File file =    new File("D:\\output.xls");
	 
	        //Create an object of FileInputStream class to read excel file
	 
	        FileInputStream inputStream = new FileInputStream(file);
	 
	         	 
	        //Find the file extension by spliting file name in substing and getting only extension name
	 
	        String fileExtensionName = fileName.substring(fileName.indexOf("."));
	 
	        //Check condition if the file is xlsx file
	 
	        Workbook guru99Workbook = null;
	        
			if(fileExtensionName.equals(".xlsx")){
	 
	        //If it is xlsx file then create object of XSSFWorkbook class
	 
	        guru99Workbook = new XSSFWorkbook(inputStream);
	 
	        }
	 
	        //Check condition if the file is xls file
	 
	        else if(fileExtensionName.equals(".xls")){
	 
	            //If it is xls file then create object of XSSFWorkbook class
	 
	            guru99Workbook = new HSSFWorkbook(inputStream);
	 
	        }
	 
	         
	 
	    //Read excel sheet by sheet name    
	 
	    HSSFSheet sheet = (HSSFSheet) guru99Workbook.getSheet(sheetName);
	 
	    //Get the current count of rows in excel file
	 
	    int rowCount = sheet.getLastRowNum()-sheet.getFirstRowNum();
	 
	    //Get the first row from the sheet
	 
	    Row row = sheet.getRow(0);
	 
	    //Create a new row and append it at last of sheet
	 
	    Row newRow = sheet.createRow(rowCount+1);
	 
	    //Create a loop over the cell of newly created Row
	 
	    for(int j = 0; j < row.getLastCellNum(); j++){
	 
	        //Fill data in row
	 
	        Cell cell = newRow.createCell(j);
	 
	        cell.setCellValue(dataToWrite);
	 
	    }
	 
	    //Close input stream
	 
	    inputStream.close();
	 
	    //Create an object of FileOutputStream class to create write data in excel file
	 
	    FileOutputStream outputStream = new FileOutputStream(file);
	 
	    //write data in the excel file
	 
	    guru99Workbook.write(outputStream);
	 
	    //close output stream
	 
	    outputStream.close();	     
	 
	    }	 
	
}
