package checker.main;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Font;
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
	        XSSFWorkbook inputWorkbook = new XSSFWorkbook(fileInput);
	        //Get first/desired sheet from the workbook
	        XSSFSheet inputSheet = inputWorkbook.getSheetAt(0);
	        
	        // output file
	        XSSFWorkbook outputWorkbook = new XSSFWorkbook();
	        XSSFSheet outputSheet = outputWorkbook.createSheet("Products_2017");
	        
	        // create first row
	        int outputRowNum = 0;
	        int topCellNum = 0;
	        Row topOutputRow = outputSheet.createRow(outputRowNum++);
	        Cell topOutputCell = topOutputRow.createCell(topCellNum++);
	        topOutputCell.setCellValue("Product");
	        topOutputCell = topOutputRow.createCell(topCellNum++);
	        topOutputCell.setCellValue("Release");
	        topOutputCell = topOutputRow.createCell(topCellNum++);
	        topOutputCell.setCellValue("Language");
	        topOutputCell = topOutputRow.createCell(topCellNum++);
	        topOutputCell.setCellValue("Adopted Athena");
	        topOutputCell = topOutputRow.createCell(topCellNum++);
	        topOutputCell.setCellValue("Link");
	        
	        //Iterate through each rows one by one
	        Iterator<Row> inputRowIterator = inputSheet.iterator();
	        int numberOfRows = 0;
	        
	        // skip the first iterator
	        if (inputRowIterator.hasNext())
	        	inputRowIterator.next();
	        
	        while (inputRowIterator.hasNext())
	        {
	            Row inputRow = inputRowIterator.next();
	            Row outputRow = outputSheet.createRow(outputRowNum++);
	            
	            int CellNum = 0;
	            
	            //For each row, iterate through all the columns
	            Iterator<Cell> inputCellIterator = inputRow.cellIterator();
	            
	            String fullUrl = "http://help.autodesk.com/view/";
	            
	            int cellIteratorNum = 0;
	            Cell outputCell;
	            
	            while (inputCellIterator.hasNext()) 
	            {
	                Cell inputCell = inputCellIterator.next();
	                
	                
	                
	                cellIteratorNum++;
	                
	                //Check the cell type and format accordingly
	                switch (inputCell.getCellType()) 
	                {
	                    case Cell.CELL_TYPE_NUMERIC:
	                        //System.out.print(cell.getNumericCellValue() + "\t");
	                        Double num = inputCell.getNumericCellValue();
	                        fullUrl = fullUrl + num.intValue() + "/";
	                        outputCell = outputRow.createCell(1);
	                        outputCell.setCellValue(num.intValue());
	                        break;
	                    case Cell.CELL_TYPE_STRING:
	                        //System.out.print(cell.getStringCellValue() + "\t");
	                    	String value = inputCell.getStringCellValue();
	                    	if (!value.equals("NA")){
	                    		fullUrl = fullUrl + inputCell.getStringCellValue() + "/";
	                    	}
	                    	if (value.equals("NA")){
	                    		outputCell = outputRow.createCell(1);
		                        outputCell.setCellValue(value);
	                    	}
	                    	if (cellIteratorNum == 1){
		                        outputCell = outputRow.createCell(0);
		                        outputCell.setCellValue(value);
	                    	} 
	                    	if (cellIteratorNum == 3){
	                    		outputCell = outputRow.createCell(2);
		                        outputCell.setCellValue(value);
	                    	}
	                        break;
	                }
	            }
	            //System.out.println("");
	            //boolean isAthena = isProductUsingAthena(driver.getPageSource());
	            
//	            if (!isAthena){
//	            	System.out.println(driver.getPageSource());
//	            }
	            //System.out.println(isAthena + " - " + fullUrl);
	            
	            driver.get(fullUrl);
	            Thread.sleep(2000);
	            
	            boolean isAthena = isProductUsingAthena(driver.getPageSource());
	            
	            if (isAthena){
	            	outputCell = outputRow.createCell(3);
                    outputCell.setCellValue("YES");
	            	XSSFCellStyle style = outputWorkbook.createCellStyle();
	            	style.setFillForegroundColor(HSSFColor.AQUA.index);
	            	style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);

	            	XSSFFont font = outputWorkbook.createFont();
	            	font.setColor(HSSFColor.DARK_BLUE.index);
	            	font.setBoldweight(Font.BOLDWEIGHT_BOLD);
	            	style.setFont(font);
                    outputCell.setCellStyle(style);
	            } else {
	            	outputCell = outputRow.createCell(3);
                    outputCell.setCellValue("NO");
                    
	            	XSSFCellStyle style = outputWorkbook.createCellStyle();
	            	style.setFillForegroundColor(HSSFColor.LIGHT_GREEN.index);
	            	style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);

	            	XSSFFont font = outputWorkbook.createFont();
	            	font.setColor(HSSFColor.RED.index);
	            	font.setBoldweight(Font.BOLDWEIGHT_BOLD);
	            	style.setFont(font);
	            	outputCell.setCellStyle(style);
	            }
	            
	            outputCell = outputRow.createCell(4);
                outputCell.setCellValue(fullUrl);
                
	            System.out.println(isAthena + " - " + fullUrl);
	            
	            numberOfRows++;
	        }
	        System.out.println("Number of rows: " + numberOfRows);
	        
	        
	        // close input
	        fileInput.close();
	        
	        try {
	            FileOutputStream out = 
	                    new FileOutputStream(new File("Products_2017_done.xlsx"));
	            outputWorkbook.write(out);
	            out.close();
	            System.out.println("Excel written successfully..");
	             
	        } catch (FileNotFoundException e) {
	            e.printStackTrace();
	        } catch (IOException e) {
	            e.printStackTrace();
	        }
	        
	        // quit driver
	        driver.quit(); 
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
