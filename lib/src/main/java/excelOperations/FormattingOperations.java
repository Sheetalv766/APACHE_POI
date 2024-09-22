package excelOperations;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.hssf.record.DBCellRecord;
import org.apache.poi.ss.formula.WorkbookEvaluator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FormattingOperations {

	/**
	 * use this method to add background color to cell
	 * 
	 * @param filePath
	 * @param sheetName
	 * @param data
	 * @param rowIndex
	 * @param colIndex
	 * @throws IOException
	 */

	public void applyColorToCell(String filePath, String sheetName, int rowIndex, int colIndex,
			IndexedColors fillColor, FillPatternType fillPatternType) throws IOException {
		try {
			/* Step - 1 : Creating file object of existing excel file */
			File fileName = new File(filePath);

			System.out.println("filePath = " + filePath);

			/* Step - 2 : Creating input stream */
			FileInputStream file = new FileInputStream(fileName);

			/* Step - 3 : Creating workbook from input stream */
			XSSFWorkbook workbook = new XSSFWorkbook(file);

			System.out.println("workbook = " + workbook);

			/* Step - 4 : Reading first sheet of excel file */
			XSSFSheet sheet = workbook.getSheet(sheetName);

			System.out.println("sheet = " + sheet);

			/* Step - 5 : Get the Cell number using getRow and getCell */
			XSSFCell cellUpdate = sheet.getRow(rowIndex).getCell(colIndex);

			/* Step - 6 : Create the cell style sheet */
			XSSFCellStyle style = workbook.createCellStyle();

			/* Step - 7 : Set background color */
			style.setFillBackgroundColor(fillColor.getIndex());

			/* Step - 8 : Set fill pattern */
			style.setFillPattern(fillPatternType);

			/* Step - 9 : Apply the style to Cell */
			cellUpdate.setCellStyle(style);

			/* Step - 10 : Close input stream */
			file.close();

			/* Step - 11 : Creating output stream and writing the updated workbook */
			FileOutputStream os = new FileOutputStream(fileName);
			workbook.write(os);

			/* Step - 12 : Close the workbook and output stream */
			workbook.close();
			os.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	/**
	 * use this method to add background color to cell
	 * 
	 * @param filePath
	 * @param sheetName
	 * @param data
	 * @param rowIndex
	 * @param colIndex
	 * @throws IOException
	 */
	public void applyAlignmentToCell(String filePath, String sheetName, int rowIndex, int colIndex,
			HorizontalAlignment horizontalAlignment) throws IOException {

		/* Step - 1 : Creating file object of existing excel file */
		File fileName = new File(filePath);

		/* Step - 2 : Creating input stream */
		FileInputStream file = new FileInputStream(fileName);

		/* Step - 3 : Creating workbook from input stream */
		XSSFWorkbook workbook = new XSSFWorkbook(file);

		/* Step - 4 : Reading first sheet of excel file */
		XSSFSheet sheet = workbook.getSheet(sheetName);

		/* Step - 5 : Get the Cell number using getRow and getCell */
		XSSFCell cellUpdate = sheet.getRow(rowIndex).getCell(colIndex);

		/* Step - 6 : Create the cell style sheet */
		XSSFCellStyle style = workbook.createCellStyle();

		/* Step - 7 : Set Alignment */
		style.setAlignment(horizontalAlignment);

		/* Step - 8 : Apply the style to Cell */
		cellUpdate.setCellStyle(style);

		/* Step - 9 : Close input stream */
		file.close();

		/* Step - 10 : Creating output stream and writing the updated workbook */
		FileOutputStream os = new FileOutputStream(fileName);
		workbook.write(os);

		/* Step - 11 : Close the workbook and output stream */
		workbook.close();
		os.close();

	}

	/**
	 * use this method to add row's and apply font color into the existing excel
	 * 
	 * @param filePath
	 * @param sheetName
	 * @param dataToWrite
	 * @throws IOException
	 */
	public void applyFontToRow(String filePath, String sheetName, Object[][] dataToWrite,
			String fontName) throws IOException {

		/* Step - 1 : Creating file object of existing excel file */
		File fileName = new File(filePath);

		/* Step - 2 : Creating input stream */
		FileInputStream file = new FileInputStream(fileName);

		/* Step - 3 : Creating workbook from input stream */
		XSSFWorkbook workbook = new XSSFWorkbook(file);

		/* Step - 4 : Reading first sheet of excel file */
		XSSFSheet sheet = workbook.getSheet(sheetName);

		/* Step - 5 : Getting the last row number of existing records */
		int rowCount = sheet.getLastRowNum();

		/**
		 * Step - 6 : Iterating dataToWrite to update* a.Create new row from the next row count
		 * b.Creating new cell and setting the value
		 */

		for (Object[] data : dataToWrite) {

			// Creating new row from the next row count
			Row row = sheet.createRow(++rowCount);

			int columnCount = 0;

			// Iterating informations
			for (Object info : data) {

				// Creating new cell and setting the value
				Cell cell = row.createCell(columnCount++);
				if (info instanceof String) {
					cell.setCellValue((String) info);
				} else if (info instanceof Integer) {
					cell.setCellValue((Integer) info);
				}
			}
		}

		/* Step - 7 : Close input stream */
		file.close();

		/* Step - 8 : Create output stream and writing the updated workbook */
		FileOutputStream os = new FileOutputStream(fileName);
		workbook.write(os);

		/* Step - 9 : Close the workbook and output stream */
		workbook.close();
		os.close();
	}

	public void validateFormattingUpdates(String filePath, String sheetName) {
		// Creating file object of existing excel file
		File fileName = new File(filePath);

		try {

			FileInputStream file = new FileInputStream(fileName);

			// Create Workbook instance holding reference to .xlsx file
			XSSFWorkbook workbook = new XSSFWorkbook(file);

			// Get first/desired sheet from the workbook
			XSSFSheet sheet = workbook.getSheet(sheetName);

			/* Verify setting color to cell B2 */
			XSSFCell cell2Style = sheet.getRow(1).getCell(1);

			XSSFCellStyle style = cell2Style.getCellStyle();

			System.out.println(
					"Color of cell B2 is: " + style.getFillBackgroundColorColor().getARGBHex());

			/* Verify setting alignment of "Survey" column to right alignment */
			cell2Style = sheet.getRow(0).getCell(3);
			style = cell2Style.getCellStyle();

			System.out.println(
					"Alignment of cell A4 (Survey) is: " + style.getAlignment().toString());

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

	public void run() throws IOException {

		// Call the desired methods
		String filePath = System.getProperty("user.dir") + "/src/main/resources/Activity.xlsx";
		String worksheetName = "Country Population";

		Object[][] countryRecord = {{"Israel", "Jerusalem", "9.2", "24-02-2021", "22145"},};

		// Add “Green” background color to Cell “B2”
		this.applyColorToCell(filePath, worksheetName, 1, 1, IndexedColors.BRIGHT_GREEN,
				FillPatternType.THICK_BACKWARD_DIAG);

		// Add “Right Alignment” to Column “Survey Date”
		this.applyAlignmentToCell(filePath, worksheetName, 0, 3, HorizontalAlignment.RIGHT);

		// Add font “Verdana” while entering the new row to your worksheet
		this.applyFontToRow(filePath, worksheetName, countryRecord, "Verdana");

		/* Add your logic to make the formatting updates above this line */


		// Utility method to verify formatting updates
		this.validateFormattingUpdates(filePath, worksheetName);
	}
}


