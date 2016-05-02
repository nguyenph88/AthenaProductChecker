package checker.main;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.phantomjs.PhantomJSDriver;
import org.openqa.selenium.phantomjs.PhantomJSDriverService;
import org.openqa.selenium.remote.DesiredCapabilities;

public class Checker {

	public static void main(String[] args) {
	    try {
			// current dir
			String currentDir = System.getProperty("user.dir");
			
	    	/////// open browser
			//WebDriver driver = new FirefoxDriver();
			
			DesiredCapabilities caps = new DesiredCapabilities();
			caps.setCapability(PhantomJSDriverService.PHANTOMJS_EXECUTABLE_PATH_PROPERTY,
			                currentDir + "\\phantomjs.exe");                  
			WebDriver driver = new PhantomJSDriver(caps);
			
	        FileInputStream fileInput = new FileInputStream(new File("Products_2017.xlsx"));

	        //Create Workbook instance holding reference to .xlsx file
	        XSSFWorkbook workbook = new XSSFWorkbook(fileInput);

	        //Get first/desired sheet from the workbook
	        XSSFSheet sheet = workbook.getSheetAt(0);

	        //Iterate through each rows one by one
	        Iterator<Row> rowIterator = sheet.iterator();
	        int numberOfRows = 0;
	        
	        // skip the first iterator
	        if (rowIterator.hasNext())
	        	rowIterator.next();
	        
	        while (rowIterator.hasNext())
	        {
	            Row row = rowIterator.next();
	            //For each row, iterate through all the columns
	            Iterator<Cell> cellIterator = row.cellIterator();
	            
	            String fullUrl = "http://help.autodesk.com/view/";
	            
	            while (cellIterator.hasNext()) 
	            {
	                Cell cell = cellIterator.next();
	                //Check the cell type and format accordingly
	                switch (cell.getCellType()) 
	                {
	                    case Cell.CELL_TYPE_NUMERIC:
	                        //System.out.print(cell.getNumericCellValue() + "\t");
	                        Double num = cell.getNumericCellValue();
	                        fullUrl = fullUrl + num.intValue() + "/";
	                        break;
	                    case Cell.CELL_TYPE_STRING:
	                        //System.out.print(cell.getStringCellValue() + "\t");
	                        fullUrl = fullUrl + cell.getStringCellValue() + "/";
	                        break;
	                }
	            }
	            //System.out.println("");
	            boolean isAthena = isProductUsingAthena(driver.getPageSource());
	            System.out.println(isAthena + " - " + fullUrl);
	            
	            driver.get(fullUrl);
	            Thread.sleep(2000);
	            //System.out.println(driver.getPageSource());
	            
	            
	            numberOfRows++;
	        }
	        System.out.println("Number of rows: " + numberOfRows);
	        
	        
	        fileInput.close();
	        driver.quit(); // quit driver
	    } catch (Exception e) {
	        e.printStackTrace();
	    }
	}
	
	public static boolean isProductUsingAthena(String source){
		if (source.contains("/view/athena/modules/")){
			return true;
		}
		if (source.contains("/view/clientframework/modules/")){
			return false;
		} 
		return false;
	}
	

}
