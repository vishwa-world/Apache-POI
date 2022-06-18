package readoperationssolution;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FormatingOperations {

	/**
	 * use this method to add background color to cell
	 * @param filePath
	 * @param sheetName
	 * @param data
	 * @param rowIndex
	 * @param colIndex
	 */
	public void applyColorToCell(String filePath, String sheetName,  int rowIndex, int colIndex) {
		 
		// Creating file object of existing excel file
	     File fileName = new File(filePath);
	     
		try {

			FileInputStream file = new FileInputStream(fileName);

			// Create Workbook instance holding reference to .xlsx file
			XSSFWorkbook workbook = new XSSFWorkbook(file);

			// Get first/desired sheet from the workbook
			XSSFSheet sheet = workbook.getSheet(sheetName);

			// Get the Cell number using getRow and getCell
			XSSFCell cell2Style = sheet.getRow(rowIndex).getCell(colIndex);

			XSSFCellStyle style = workbook.createCellStyle();
			style.setFillBackgroundColor(IndexedColors.BRIGHT_GREEN.getIndex());
			style.setFillPattern(FillPatternType.THICK_BACKWARD_DIAG);
			
			cell2Style.setCellStyle(style);
			
			// Close input stream
			file.close();

			// Crating output stream and writing the updated workbook
			FileOutputStream os = new FileOutputStream(fileName);
			workbook.write(os);

			// Close the workbook and output stream
			workbook.close();
			os.close();

			System.out.println("Excel file has been updated successfully.");

		} catch (Exception e) {
			System.err.println("Exception while updating an existing excel file.");
			e.printStackTrace();
		}

	}
	
	
	/**
	 * use this method to add background color to cell
	 * @param filePath
	 * @param sheetName
	 * @param data
	 * @param rowIndex
	 * @param colIndex
	 */
	public void applyAlignmentToCell(String filePath, String sheetName, int rowIndex, int colIndex) {
		 
		// Creating file object of existing excel file
	     File fileName = new File(filePath);
	     
		try {

			FileInputStream file = new FileInputStream(fileName);

			// Create Workbook instance holding reference to .xlsx file
			XSSFWorkbook workbook = new XSSFWorkbook(file);

			// Get first/desired sheet from the workbook
			XSSFSheet sheet = workbook.getSheet(sheetName);

			// Get the Cell number using getRow and getCell
			XSSFCell cell2Style = sheet.getRow(rowIndex).getCell(colIndex);
			
			XSSFCellStyle style = workbook.createCellStyle();
			style.setAlignment(HorizontalAlignment.RIGHT);
			
			cell2Style.setCellStyle(style);
			
			// Close input stream
			file.close();

			// Crating output stream and writing the updated workbook
			FileOutputStream os = new FileOutputStream(fileName);
			workbook.write(os);

			// Close the workbook and output stream
			workbook.close();
			os.close();

			System.out.println("Excel file has been updated successfully.");

		} catch (Exception e) {
			System.err.println("Exception while updating an existing excel file.");
			e.printStackTrace();
		}

	}

	/**
	 * use this method to add row's and apply font color into the existing excel
	 * @param filePath
	 * @param sheetName
	 * @param dataToWrite
	 */
	
	public void applyFontToRow(String filePath,String sheetName,Object[][] dataToWrite) {
		 
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

	           XSSFFont font = workbook.createFont();
			   font.setFontName("Verdana");
	 
				XSSFCellStyle style = workbook.createCellStyle();
				style.setFont(font);
				
	         //Iterating new students to update
	         for (Object[] data : dataToWrite) {
	              
	             //Creating new row from the next row count
	        	 XSSFRow row = sheet.createRow(++rowCount);
	        	 int columnCount = 0;
	             //Iterating student informations
	             for (Object info : data) {
	                 //Creating new cell and setting the value
	            	 XSSFCell cell = row.createCell(columnCount++);
	            	 cell.setCellStyle(style);
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
	public static void main(String[] args) {
		FormatingOperations formatingOperations = new FormatingOperations();
		
		
		 //New students records to update in excel file
	     Object[][] countryRecord = {
	             {"Israel","Jerusalem","9.2","24-02-2021","22145"},
	     };
	     formatingOperations.applyColorToCell(System.getProperty("user.dir") + "\\src\\test\\resources\\CountryInfo.xlsx",
					"Country Population",1,3);
	}

}
