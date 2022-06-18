package readoperationssolution;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteOpetaions {
	
	/**
	 * use this method to add row's into the existing excel
	 * @param filePath
	 * @param sheetName
	 * @param dataToWrite
	 */
	 public void writeInToExcel(String filePath,String sheetName,Object[][] dataToWrite) {
		 
		 // Creating file object of existing excel file
	     File fileName = new File(filePath);
	      
	     try {
	    	 
	 		FileInputStream file = new FileInputStream(fileName);

	 		// Create Workbook instance holding reference to .xlsx file
	 		XSSFWorkbook workbook = new XSSFWorkbook(file);

	 		// Get first/desired sheet from the workbook
	 		XSSFSheet sheet = workbook.getSheet(sheetName);
	 		
	         //Getting the count of existing records
	         int rowCount = sheet.getLastRowNum();

	         //Iterating new students to update
	         for (Object[] data : dataToWrite) {
	              
	             //Creating new row from the next row count
	        	 XSSFRow row = sheet.createRow(++rowCount);
	        	 int columnCount = 0;
	             //Iterating student informations
	             for (Object info : data) {
	                 //Creating new cell and setting the value
	            	 XSSFCell cell = row.createCell(columnCount++);
	                 if (info instanceof String) {
	                     cell.setCellValue((String) info);
	                 } else if (info instanceof Integer) {
	                     cell.setCellValue((Integer) info);
	                 }
	                 else if (info instanceof Double) {
	                     cell.setCellValue((Double) info);
	                 }
	             }
	         }
	         //Close input stream
	         file.close();

	         //Crating output stream and writing the updated workbook
	         FileOutputStream os = new FileOutputStream(fileName);
	         workbook.write(os);
	          
	         //Close the workbook and output stream
	         workbook.close();
	         os.close();
	          
	         System.out.println("Excel file has been updated successfully.");
	          
	     } catch (Exception e) {
	         System.err.println("Exception while updating an existing excel file.");
	         e.printStackTrace();
	     }
	 }
	 
	 /**
	  * use this method to update the existing value of Cell
	  * @param filePath
	  * @param sheetName
	  * @param data
	  * @param rowIndex
	  * @param colIndex
	  */
	 public void updateCellValue(String filePath,String sheetName,String data,int rowIndex,int colIndex) {
		 
		 // Creating file object of existing excel file
	     File fileName = new File(filePath);
	      
	     try {
	    	 
	 		FileInputStream file = new FileInputStream(fileName);

	 		// Create Workbook instance holding reference to .xlsx file
	 		XSSFWorkbook workbook = new XSSFWorkbook(file);

	 		// Get first/desired sheet from the workbook
	 		XSSFSheet sheet = workbook.getSheet(sheetName);
	 		
	 		//Get the Cell number using getRow and getCell
	 		XSSFCell cell2Update = sheet.getRow(rowIndex).getCell(colIndex);
	 		
	 		cell2Update.setCellValue(data);
	 		
	 		//Close input stream
	         file.close();

	         //Crating output stream and writing the updated workbook
	         FileOutputStream os = new FileOutputStream(fileName);
	         workbook.write(os);
	          
	         //Close the workbook and output stream
	         workbook.close();
	         os.close();
	          
	         System.out.println("Excel file has been updated successfully.");
	 		
	     }catch (Exception e) {
	         System.err.println("Exception while updating an existing excel file.");
	         e.printStackTrace();
	     }
	 }
	 
	 public static void main(String[] args) {
		 
		 //New students records to update in excel file
	     Object[][] countryRecord = {
	             {"UK","London","6.72","15-02-2021"},
	             {"US","Washington,D.C","32.95","09-02-2021"}
	     };
	     
		 WriteOpetaions wp = new WriteOpetaions();
		 wp.updateCellValue(System.getProperty("user.dir")+"\\src\\test\\resources\\CountryInfo.xlsx", "Country Population","27-02-2021",1,3);
		 
	}

}
