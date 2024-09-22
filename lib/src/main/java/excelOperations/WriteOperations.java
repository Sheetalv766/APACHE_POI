package excelOperations;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.interactions.InputSource;

public class WriteOperations {

	/**
	 * use this method to add row's into the existing excel
	 * 
	 * @param filePath
	 * @param sheetName
	 * @param dataToWrite
	 * @throws IOException
	 */
	public void writeInToExcel(String filePath, String sheetName, Object[][] dataToWrite)
			throws IOException {
		System.out.println("**Add new row to: " + sheetName + "**");

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

		System.out.println("Excel file UPDATED successfully.\n");

	}

	/**
	 * use this method to update the particular Cell value
	 * 
	 * @param filePath
	 * @param sheetName
	 * @param data
	 * @param rowIndex
	 * @param colIndex
	 * @throws IOException
	 */
	public void updateCellValue(String filePath, String sheetName, String data, int rowIndex,
			int colIndex) throws IOException {

		/* Step - 1 : Creating file object of existing excel file */
		File fileName = new File(filePath);

		/* Step - 2 : Creating input stream */
		FileInputStream file = new FileInputStream(fileName);

		/* Step - 3 : Creating workbook from input stream */
		XSSFWorkbook workbook = new XSSFWorkbook(file);

		/* Step - 4 : Reading first sheet of excel file */
		//sheet --> object
		XSSFSheet sheet = workbook.getSheet(sheetName);

		/* Step - 5 : Get the Cell number using getRow and getCell */
		XSSFCell cellUpdate = sheet.getRow(rowIndex).getCell(colIndex);

		/* Step - 6 : Update the cell */
		cellUpdate.setCellValue(data);

		/* Step - 7 : Close input stream */
		file.close();

		/* Step - 8 : Creating output stream and writing the updated workbook */
		FileOutputStream os = new FileOutputStream(fileName);
		workbook.write(os);

		/* Step - 9 : Close the workbook and output stream */
		workbook.close();
		os.close();

		System.out.println("Cell value updated successfully.\n");

	}

	public void addColumn(String filePath, String sheetName, String[] colValues)
			throws IOException {

		System.out.println("**Add new column to sheet: " + sheetName + "**");

		/* Step - 1 : Creating file object of existing excel file */
		File fileName = new File(filePath);

		/* Step - 2 : Creating input stream */
		FileInputStream file = new FileInputStream(fileName);

		/* Step - 3 : Creating workbook from input stream */
		XSSFWorkbook workbook = new XSSFWorkbook(file);

		/* Step - 4 : Reading first sheet of excel file */
		XSSFSheet sheet = workbook.getSheet(sheetName);

		/* Step - 5 : Get all the rows and add a new cell to it at the end */
		Iterator<Row> iterator = sheet.iterator();

		/* Step - 6 : Close input stream */
		file.close();

		/* Step - 7 : Creating output stream and writing the updated workbook */
		FileOutputStream os = new FileOutputStream(fileName);
		workbook.write(os);

		/* Step - 8 : Close the workbook and output stream */
		workbook.close();
		os.close();

		System.out.println("Excel file UPDATED");

	}

	public void run() throws IOException {
		// Call the desired methods
		String filePath = System.getProperty("user.dir") + "/src/main/resources/Activity.xlsx";
		String worksheetName = "Country Population";

		Object[][] countryPop = {{"UK", "London", "6.72", "15-02-2021"},
				{"US", "Washington,D.C", "32.95", "09-02-2021"}};

		// 1. Add the following rows into existing worksheet “Country Population”
		this.writeInToExcel(filePath, worksheetName, countryPop);

		// 2. Update the “Survey Date” column for country “India” from “27-02-2011” to “27-02-2021”
		this.updateCellValue(filePath, worksheetName, "27-02-2021", 1, 3);

		// 3. Create a new column “Area(Km)”
		String[] colValues =
				{"Area (Km2)", "3287000", "54394", "30688", "302068", "17100000", "42933"};
		this.addColumn(filePath, worksheetName, colValues);
	}
}
