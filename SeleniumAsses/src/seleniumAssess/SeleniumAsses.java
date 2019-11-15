package seleniumAssess;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class SeleniumAsses {
	public static String vSearch;
	public static int xlRows, xlCols;
	public static String xData[][];

	public static void main(String[] args) throws Exception {

		getDataFromTheSheet();
		for (int i = 1; i < xlRows; i++) {
			if (xData[i][1].equals("Y")) {

				// Gets Data from the Excel Sheet to search
				vSearch = getDataFromVariable(i);

				// Gets Search Result from the WebDriver
				String result = webDriverFunctions();

				// Sends the Result to the 2D variable
				setDataToVariable(result, i);

			}
		}
		sendDatatoTheSheet();
	}

	// Get the value from the cell
	public static String cellToString(HSSFCell cell) {
		int type = cell.getCellType();
		Object result;
		switch (type) {
		case HSSFCell.CELL_TYPE_NUMERIC:
			result = cell.getNumericCellValue();
			break;
		case HSSFCell.CELL_TYPE_STRING:
			result = cell.getStringCellValue();
			break;
		case HSSFCell.CELL_TYPE_FORMULA:
			throw new RuntimeException("We cannot evaluate formula");
		case HSSFCell.CELL_TYPE_BLANK:
			result = "-";
		case HSSFCell.CELL_TYPE_BOOLEAN:
			result = cell.getBooleanCellValue();
		case HSSFCell.CELL_TYPE_ERROR:
			result = "This cell has some error";
		default:
			throw new RuntimeException("We do not support this cell type");
		}
		return result.toString();

	}

	// To read the function from the Sheet
	public static void xlRead(String sPath) throws Exception {
		File myFile = new File(sPath);
		FileInputStream myStream = new FileInputStream(myFile);
		HSSFWorkbook myworkbook = new HSSFWorkbook(myStream);
		HSSFSheet mySheet = myworkbook.getSheetAt(0);
		xlRows = mySheet.getLastRowNum() + 1;
		xlCols = mySheet.getRow(0).getLastCellNum();
		xData = new String[xlRows][xlCols];
		for (int i = 0; i < xlRows; i++) {
			HSSFRow row = mySheet.getRow(i);
			for (short j = 0; j < xlCols; j++) {
				HSSFCell cell = row.getCell(j);
				String value = cellToString(cell);
				xData[i][j] = value;
				System.out.print("-" + xData[i][j]);
			}
			System.out.println();
		}
	}

	// Writes the value to the Excel sheet
	public static void xlwrite(String xlpath1, String[][] xData) throws Exception {
		System.out.println("Inside XL Write");
		File myFile1 = new File(xlpath1);
		FileOutputStream fout = new FileOutputStream(myFile1);
		HSSFWorkbook wb = new HSSFWorkbook();
		HSSFSheet mySheet1 = wb.createSheet("TestResults");
		for (int i = 0; i < xlRows; i++) {
			HSSFRow row1 = mySheet1.createRow(i);
			for (short j = 0; j < xlCols; j++) {
				HSSFCell cell1 = row1.createCell(j);
				cell1.setCellType(HSSFCell.CELL_TYPE_STRING);
				cell1.setCellValue(xData[i][j]);
			}
		}
		wb.write(fout);
		fout.flush();
		fout.close();
	}

	// To perform the search function to get the title
	public static String webDriverFunctions() throws InterruptedException {
		System.setProperty("webdriver.chrome.driver",
				"C:\\Users\\gparthasara2\\Downloads\\chromedriver_win32\\chromedriver.exe");
		WebDriver myDriverChrome = new ChromeDriver();
		myDriverChrome.manage().window().maximize();
		myDriverChrome.get("https://www.yahoo.com");
		Thread.sleep(1000);
		myDriverChrome.findElement(By.name("p")).sendKeys(vSearch, Keys.ENTER);
		Thread.sleep(1000);
		String result = myDriverChrome.getTitle();
		myDriverChrome.close();
		return result;
	}

	// Get data from the 2D Variable
	public static String getDataFromVariable(int i) {
		return xData[i][0];
	}

	// Sends result to the 2D String variable
	public static void setDataToVariable(String s, int i) {
		xData[i][2] = s;
	}

	// Get Data from the Sheet to the 2D Variable
	public static void getDataFromTheSheet() throws Exception {
		xlRead("C:\\Users\\gparthasara2\\Desktop\\Folders\\Training\\Testing\\demo2\\YahooDDF.xls");

	}

	// Send Data to the Sheet
	public static void sendDatatoTheSheet() throws Exception {
		xlwrite("C:\\Users\\gparthasara2\\Desktop\\Folders\\Training\\Testing\\demo2\\YahooDDF.xls", xData);

	}

}
