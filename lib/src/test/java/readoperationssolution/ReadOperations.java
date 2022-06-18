package readoperationssolution;

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

	public static void main(String[] args) {
		ReadOperations readOperations = new ReadOperations();
		try {
			readOperations.getColunmValue(System.getProperty("user.dir")+"\\src\\test\\resources\\CountryInfo.xlsx", "Country Population", 1);
			readOperations.getRowValue(System.getProperty("user.dir")+"\\src\\test\\resources\\CountryInfo.xlsx", "Country Population", 1);
			readOperations.getCellValue(System.getProperty("user.dir")+"\\src\\test\\resources\\CountryInfo.xlsx", "Country Population", 3,1);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
