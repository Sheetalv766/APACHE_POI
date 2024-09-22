package excelOperations;

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
	 * 
	 * @param filePath
	 * @param sheetName
	 * @throws IOException
	 */
	public void readCompleteExcel(String filePath, String sheetName) throws IOException {

		/*
		 * Step - 1 : Read the excel file using FileInputStream to obtain input bytes from a file.
		 * a. Create the object of File b. Create the object of FileInputStream
		 */
		System.out.println("*Printing complete worksheet data of: " + sheetName + "*\n");

		File fileName = new File(filePath);
		FileInputStream file = new FileInputStream(fileName);

		/* Step - 2 : Create Workbook instance holding reference to .xlsx file */

		XSSFWorkbook workbook = new XSSFWorkbook(file);

		/*
		 * Step - 3 : Get first/desired sheet from the workbook
		 */

		XSSFSheet sheet = workbook.getSheet(sheetName);

		/* Step - 4 : Get the last row number */

		int rowsCount = sheet.getLastRowNum();

		/*
		 * Step - 5 : Get the last cell number
		 */
		// only give the count till the last cell number which have value
		int colsCount = sheet.getRow(1).getLastCellNum();

		/*
		 * Step - 6 : Use a for each loop to iterate the row a. get the row b. using for each loop
		 * iterate over Cell of the row c. using switch statement check Cell type d. print the cell
		 * value
		 *
		 */

		// outer for loop to iterate each row
		for (int outer = 0; outer <= colsCount; outer++) {
			XSSFRow rows = sheet.getRow(outer);
			// inner for loop to iterate each cell
			for (int inner = 0; inner < colsCount; inner++) {
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
	 * 
	 * @param filePath
	 * @param sheetName
	 * @param rowIndex
	 * @throws IOException
	 */
	public void getRowValue(String filePath, String sheetName, int rowIndex) throws IOException {
		/*
		 * Step - 1 : Read the excel file using FileInputStream to obtain input bytes from a file.
		 * a. Create the object of File b. Create the object of FileInputStream
		 */
		System.out.println("*Print data in row no. " + rowIndex + " : " + sheetName + "*\n");

		File fileName = new File(filePath);
		FileInputStream file = new FileInputStream(fileName);

		/* Step - 2 : Create Workbook instance holding reference to .xlsx file */

		XSSFWorkbook workbook = new XSSFWorkbook(file);

		/*
		 * Step - 3 : Get first/desired sheet from the workbook
		 */
		XSSFSheet sheet = workbook.getSheet(sheetName);

		/* Step - 4 : Get the desire row */

		XSSFRow row = sheet.getRow(rowIndex);

		/*
		 * Step - 5 : Iterate over over each Cell using for each loop a. using switch statement
		 * check Cell type b. print the cell value
		 */

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
	}

	/**
	 * use this method to read column value
	 *
	 * @param filePath
	 * @param sheetName
	 * @param columnIndex
	 * @throws IOException
	 */
	public void getColunmValue(String filePath, String sheetName, int columnIndex)
			throws IOException {

		/*
		 * Step - 1 : Read the excel file using FileInputStream to obtain input bytes from a file.
		 * a. Create the object of File b. Create the object of FileInputStream
		 */

		System.out
				.println("*Print data in col no. " + columnIndex + " : " + sheetName + "*\n");

		File fileName = new File(filePath);
		FileInputStream file = new FileInputStream(fileName);

		/* Step - 2 : Create Workbook instance holding reference to .xlsx file */
		XSSFWorkbook workbook = new XSSFWorkbook(file);

		/*
		 * Step - 3 : Get first/desired sheet from the workbook
		 */

		XSSFSheet sheet = workbook.getSheet(sheetName);

		/*
		 * Step - 4 : Using for each loop iterate over each row a. using for each loop iterate over
		 * each Cell of row b. Compare the column index for which you want print the values. if,
		 * match found print the cell value
		 */

		for (Row row : sheet) {
			// Using foreach loop iterate over each Cell of row
			for (Cell cell : row) {
				// Compare the column index for which you want print the values
				// if match found print the cell value
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
	 * 
	 * @param filePath
	 * @param sheetName
	 * @param rowIndex
	 * @param colIndex
	 * @throws IOException
	 */
	public void getCellValue(String filePath, String sheetName, int rowIndex, int colIndex)
			throws IOException {

		/*
		 * Step - 1 : Read the excel file using FileInputStream to obtain input bytes from a file.
		 * a. Create the object of File b. Create the object of FileInputStream
		 */
		System.out.println("*Printing data row no. " + rowIndex + " col no. " + colIndex + " of: "
				+ sheetName + "*\n");

		File fileName = new File(filePath);
		FileInputStream file = new FileInputStream(fileName);

		/* Step - 2 : Create Workbook instance holding reference to .xlsx file */

		XSSFWorkbook workbook = new XSSFWorkbook(file);

		/*
		 * Step - 3 : Get first/desired sheet from the workbook
		 */

		XSSFSheet sheet = workbook.getSheet(sheetName);
		/*
		 * Step - 4 : Get the row from which you want to read the cell data
		 */

		XSSFRow row = sheet.getRow(rowIndex);

		/* Step - 5 : Get the Cell value by passing the column index */

		XSSFCell cell = row.getCell(colIndex);

		/* Step - 6 : Print the cell value */

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
	}

	public void run() {
		// Call the desired methods
		String filePath = System.getProperty("user.dir") + "/src/main/resources/Activity.xlsx";
		String worksheetName = "Country Population";
		try {
			this.readCompleteExcel(filePath, worksheetName);
			this.getRowValue(filePath, worksheetName, 1);
			this.getColunmValue(filePath, worksheetName, 1);
			this.getCellValue(filePath, worksheetName, 3, 1);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
