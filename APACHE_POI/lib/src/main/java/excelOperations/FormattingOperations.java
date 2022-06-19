package excelOperations;

import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;

public class FormattingOperations {

	/**
	 * use this method to add background color to cell
	 * 
	 * @param filePath
	 * @param sheetName
	 * @param data
	 * @param rowIndex
	 * @param colIndex
	 */

	public void applyColorToCell(String filePath, String sheetName, int rowIndex, int colIndex,
			IndexedColors fillColor, FillPatternType fillPatternType) {
		/* Step - 1 : Creating file object of existing excel file */

		/* Step - 2 : Creating input stream */

		/* Step - 3 : Creating workbook from input stream */

		/* Step - 4 : Reading first sheet of excel file */

		/* Step - 5 : Get the Cell number using getRow and getCell */

		/* Step - 6 : Create the cell style sheet */

		/* Step - 7 : Set background color */

		/* Step - 8 : Set fill pattern */

		/* Step - 9 : Apply the style to Cell */

		/* Step - 10 : Close input stream */

		/* Step - 11 : Creating output stream and writing the updated workbook */

		/* Step - 12 : Close the workbook and output stream */

	}

	/**
	 * use this method to add background color to cell
	 * 
	 * @param filePath
	 * @param sheetName
	 * @param data
	 * @param rowIndex
	 * @param colIndex
	 */
	public void applyAlignmentToCell(String filePath, String sheetName, int rowIndex, int colIndex,
			HorizontalAlignment horizontalAlignment) {

		/* Step - 1 : Creating file object of existing excel file */

		/* Step - 2 : Creating input stream */

		/* Step - 3 : Creating workbook from input stream */

		/* Step - 4 : Reading first sheet of excel file */

		/* Step - 5 : Get the Cell number using getRow and getCell */

		/* Step - 6 : Create the cell style sheet */

		/* Step - 7 : Set Alignment */

		/* Step - 8 : Apply the style to Cell */

		/* Step - 9 : Close input stream */

		/* Step - 10 : Creating output stream and writing the updated workbook */

		/* Step - 11 : Close the workbook and output stream */

	}

	/**
	 * use this method to add row's and apply font color into the existing excel
	 * 
	 * @param filePath
	 * @param sheetName
	 * @param dataToWrite
	 */
	public void applyFontToRow(String filePath, String sheetName, Object[][] dataToWrite,
			String fontName) {

		/* Step - 1 : Creating file object of existing excel file */

		/* Step - 2 : Creating input stream */

		/* Step - 3 : Creating workbook from input stream */

		/* Step - 4 : Reading first sheet of excel file */

		/* Step - 5 : Getting the last row number of existing records */

		/**
		 * Step - 6 : Iterating dataToWrite to update* a.Create new row from the next row count
		 * b.Creating new cell and setting the value
		 */

		/* Step - 7 : Close input stream */

		/* Step - 8 : Create output stream and writing the updated workbook */

		/* Step - 9 : Close the workbook and output stream */
	}

	public void run() {
		// Call the desired methods
	}
}
