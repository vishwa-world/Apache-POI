// package readoperations;

// import java.io.IOException;

// public class ReadOperations {
	
// 	/**
// 	 * use this method to read the complete excel file
// 	 * @param filePath
// 	 * @param sheetName
// 	 * @throws IOException
// 	 */
// 	public void readCompleteExcel(String filePath,String sheetName) throws IOException {
		
// 		/*
// 		 * Step - 1 : Read the excel file using FileInputStream to obtain input bytes
// 		 * from a file. 
// 		 * a. Create the object of File 
// 		 * b. Create the object of FileInputStream
// 		 */	
		
		
// 		/* Step - 2 : Create Workbook instance holding reference to .xlsx file */
		

// 		/*
// 		 * Step - 3 : Get first/desired sheet from the workbook
// 		 */	
		
// 		/* Step - 4 : Get the last row number */
		
// 		/*
// 		 * Step - 5 : Get the last cell number
// 		 */
		
// 		/* Step - 6 : Use a for each loop to iterate the row 
// 		 *   a. get the row 
// 		 *   b. using for each loop iterate over Cell of the row
// 		 *   c. using switch statement check Cell type
// 		 *   d. print the cell value
// 		 * 
// 		 * */
// 	}
	
// 	/**
// 	 * use this method to read the row values from excel
// 	 * @param filePath
// 	 * @param sheetName
// 	 * @param rowIndex
// 	 * @throws IOException
// 	 */
// 	public void getRowValue(String filePath, String sheetName, int rowIndex) throws IOException {
// 		/*
// 		 * Step - 1 : Read the excel file using FileInputStream to obtain input bytes
// 		 * from a file. 
// 		 * a. Create the object of File 
// 		 * b. Create the object of FileInputStream
// 		 */	
		
		
// 		/* Step - 2 : Create Workbook instance holding reference to .xlsx file */
	
// 		/*
// 		 * Step - 3 : Get first/desired sheet from the workbook
// 		 */
		
// 		/* Step - 4 : Get the desire row */
		
// 		/*
// 		 * Step - 5 : Iterate over over each Cell using for each loop
// 		 * a. using switch statement check Cell type
// 		 * b. print the cell value
// 		 */
// 	}
	
// 	/**
// 	 * use this method to read column value
// 	 * 
// 	 * @param filePath
// 	 * @param sheetName
// 	 * @param columnIndex
// 	 * @throws IOException
// 	 */
// 	public void getColunmValue(String filePath,String sheetName, int columnIndex) throws IOException{
		
// 		/*
// 		 * Step - 1 : Read the excel file using FileInputStream to obtain input bytes
// 		 * from a file. 
// 		 * a. Create the object of File 
// 		 * b. Create the object of FileInputStream
// 		 */	
		
		
// 		/* Step - 2 : Create Workbook instance holding reference to .xlsx file */
	
// 		/*
// 		 * Step - 3 : Get first/desired sheet from the workbook
// 		 */
		
// 		/* Step - 4 : Using for each loop iterate over each row 
// 		 *   a. using for each loop iterate over each Cell of row
// 		 *   b. Compare the column index for which you want print the values. if, match found 
// 		 *      print the cell value 
// 		 * */
// 	}
	
// 	/**
// 	 * use this method to read a particular Cell value
// 	 * @param filePath
// 	 * @param sheetName
// 	 * @param rowIndex
// 	 * @param colIndex
// 	 * @throws IOException
// 	 */
// 	public void getCellValue(String filePath,String sheetName, int rowIndex, int colIndex) throws IOException {
		
// 		/*
// 		 * Step - 1 : Read the excel file using FileInputStream to obtain input bytes
// 		 * from a file. 
// 		 * a. Create the object of File 
// 		 * b. Create the object of FileInputStream
// 		 */	
		
// 		/* Step - 2 : Create Workbook instance holding reference to .xlsx file */
	
// 		/*
// 		 * Step - 3 : Get first/desired sheet from the workbook
// 		 */
		
// 		/*
// 		 * Step - 4 : Get the row from which you want to read the cell data
// 		 */
		
// 		/* Step - 5 : Get the Cell value by passing the column index */
		
// 		/* Step - 6 : Print the cell value */
// 	}
	
// 	public static void main(String[] args) {
// 		ReadOperations readOperations = new ReadOperations();
		
// 		//Call the desire method
// 	}
// }
package readoperations;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadOperations {

	/**
	 * use this method to read the complete excel file
	 * @param filePath
	 * @param sheetName
	 * @throws IOException
	 */
	public void readCompleteExcel(String filePath, String sheetName) throws IOException {
		File fileName = new File(filePath);

		FileInputStream file = new FileInputStream(fileName);

		// Create Workbook instance holding reference to .xlsx file
		XSSFWorkbook workbook = new XSSFWorkbook(file);

		// Get first/desired sheet from the workbook
		XSSFSheet sheet = workbook.getSheet(sheetName);
		int rowCount = sheet.getLastRowNum();

		int colsCount = sheet.getRow(1).getLastCellNum();
		for (int outer = 0; outer <= rowCount; outer++) { // outer for loop to iterate each row
			XSSFRow rows = sheet.getRow(outer);
			for (int inner = 0; inner < colsCount; inner++) { // inner for loop to iterate each cell
				XSSFCell cell = rows.getCell(inner);
				switch (cell.getCellType()) {
				case STRING:
					System.out.print(cell.getStringCellValue());
					break;
				case NUMERIC:
					System.out.print(cell.getNumericCellValue());
					break;
				case BOOLEAN:
					System.out.print(cell.getBooleanCellValue());
					break;
				default:
					break;
				}
				System.out.print(" | ");
			}
			System.out.println();
		}
	}

	/**
	 * use this method to read the row values from excel
	 * @param filePath
	 * @param sheetName
	 * @param rowIndex
	 * @throws IOException
	 */
	
	public void getRowValue(String filePath, String sheetName, int rowIndex) throws IOException {

		File fileName = new File(filePath);

		FileInputStream file = new FileInputStream(fileName);

		// Create Workbook instance holding reference to .xlsx file
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		XSSFSheet sheet = workbook.getSheet(sheetName);
		XSSFRow row = sheet.getRow(rowIndex);
		System.out.print(rowIndex + " Row values: "+ "|");
		for (Cell cell : row) {
			switch (cell.getCellType()) {
			case STRING:
				System.out.print(cell.getStringCellValue() + "|");
				break;
			case NUMERIC:
				System.out.print(cell.getNumericCellValue() + "|");
				break;
			case BOOLEAN:
				System.out.print(cell.getBooleanCellValue() + "|");
				break;
			default:
				break;
			}
		}
		System.out.println("");
	}
	
	/**
	 * use this method to read column value
	 * 
	 * @param filePath
	 * @param sheetName
	 * @param columnIndex
	 * @throws IOException
	 */
	public void getColunmValue(String filePath,String sheetName, int columnIndex) throws IOException {
		File fileName = new File(filePath);

		FileInputStream file = new FileInputStream(fileName);
		
		// Create Workbook instance holding reference to .xlsx file
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		XSSFSheet sheet = workbook.getSheet(sheetName);
		System.out.println("Printing the colunm values : ");
		for (Row row : sheet) {
			for (Cell cell : row) {
				if (cell.getColumnIndex() == columnIndex) {
					switch (cell.getCellType()) {
					case STRING:
						System.out.println("|" + cell.getStringCellValue() + "|");
						break;
					case NUMERIC:
						System.out.println("|" + cell.getNumericCellValue() + "|");
						break;
					case BOOLEAN:
						System.out.println("|" + cell.getBooleanCellValue() + "|");
						break;
					default:
						break;
					}
				}
			}
		}
	}

	/**
	 * use this method to read a particular Cell value
	 * @param filePath
	 * @param sheetName
	 * @param rowIndex
	 * @param colIndex
	 * @throws IOException
	 */
	public void getCellValue(String filePath,String sheetName, int rowIndex, int colIndex) throws IOException {
		
		File fileName = new File(filePath);
		FileInputStream file = new FileInputStream(fileName);
	
		// Create Workbook instance holding reference to .xlsx file
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		XSSFSheet sheet = workbook.getSheet(sheetName);
		XSSFRow row = sheet.getRow(rowIndex);
		XSSFCell cell = row.getCell(colIndex);
		System.out.println("Cell at value at "+rowIndex+":"+colIndex +" is :");
		switch (cell.getCellType()) {
		case STRING:
			System.out.print(cell.getStringCellValue());
			break;
		case NUMERIC:
			System.out.print(cell.getNumericCellValue());
			break;
		case BOOLEAN:
			System.out.print(cell.getBooleanCellValue());
			break;
		default:
			break;
		}
		System.out.println("");

	}

	public void run() {
		try {
			this.getColunmValue(System.getProperty("user.dir")+"\\src\\test\\resources\\CountryInfo.xlsx", "Country Population", 1);
			this.getRowValue(System.getProperty("user.dir")+"\\src\\test\\resources\\CountryInfo.xlsx", "Country Population", 1);
			this.getCellValue(System.getProperty("user.dir")+"\\src\\test\\resources\\CountryInfo.xlsx", "Country Population", 3,1);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
