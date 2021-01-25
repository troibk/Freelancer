package selenium;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.LinkedList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;

public class Tests_lifelabs_com {
	static WebDriver driver;
	List<String[]> datalist = new ArrayList<>();
	String fileName = "tests.lifelabs.com.xlsx";
	String fileResult = "tests.lifelabs.com_result.xlsx";

	public Tests_lifelabs_com() {
		System.setProperty("webdriver.chrome.driver", "D:\\Freelancer\\thoan_excel\\chromedriver\\chromedriver.exe");
		driver = new ChromeDriver();
		driver.manage().window().maximize();
	}

	public static void main(String[] args) {

		String url = "http://tests.lifelabs.com/Laboratory_Test_Information/Browse_By_Test.aspx";
		Tests_lifelabs_com t = new Tests_lifelabs_com();
//		 t.processWeb(url);
		t.processExcel();

	}

	public void processWeb(String url) {
		driver.get(url);
		try {
			Thread.sleep(3000);
			List<WebElement> items = driver.findElement(By.id("resultsDiv")).findElements(By.tagName("a"));
			for (int i=1; i<items.size();i++) {
				String link = items.get(i).getAttribute("href");
				datalist.add(new String[] { link, "LifeLabs","All Regions","Ontario" });
			}
			appendExcelFile(fileName, "");
		} catch (Exception e) {
			e.printStackTrace();
		}

		System.out.println("===============Finish====================");
	}

	public void processExcel() {

		try {

			FileInputStream inputFile = new FileInputStream(
					new File("D:\\Freelancer\\thoan_excel\\Results\\" + fileName));
			XSSFWorkbook wb = new XSSFWorkbook(inputFile);
			XSSFSheet table1Sheet = wb.getSheetAt(0);

			if (table1Sheet == null) {
				System.out.println("KO CO SHEET :" + table1Sheet);
			} else {

				for (int i = 4900; i < 5000; i++) {
					try {

						Cell cell = table1Sheet.getRow(i).getCell(0);
						if (cell == null) {
							continue;
						}

						String url = table1Sheet.getRow(i).getCell(0).getStringCellValue();
						String company = table1Sheet.getRow(i).getCell(1).getStringCellValue();
						String city = table1Sheet.getRow(i).getCell(2).getStringCellValue();
						String country = table1Sheet.getRow(i).getCell(3).getStringCellValue();
						boolean isNext = processPage(url, company, city, country, i + 1);
						if (datalist.size()==100) {
							appendExcelFile(fileResult, "");
							Thread.sleep(3000);
							System.out.println("=========="+(i+1));
						}
						if(!isNext) {
							break;
						}
						
					} catch (Exception e) {
						System.out.println("EEEEEERRRRRRRRRR: " + i);
						e.printStackTrace();
					}
				}
				appendExcelFile(fileResult, "");
				System.out.println("===============Finish====================");
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public boolean processPage(String url, String company, String city, String country, int rowIndex) {

		try {
			driver.get(url);
			Thread.sleep(2000);
			
			String name="";
			String component="";
			WebElement linkE=driver.findElement(By.xpath("//*[@id=\"ctl00_ctl00_ctl00_bodyMain_cellOne_MainCell_bodyTabContainer\"]/div[4]/div[3]/div[2]/ul/li[1]/a"));
			clickAction(linkE);
			Thread.sleep(500);
			String link = driver.getCurrentUrl();
			List<WebElement> nameE = driver.findElements(By.cssSelector("div[class='testTitle']"));
			if(!nameE.isEmpty()) {
				name=nameE.get(0).getText();
			}
			
			List<WebElement> componentE = driver.findElements(By.cssSelector("div[class='testaka']"));
			if(!componentE.isEmpty()) {
				component=componentE.get(0).getText();
			}
			WebElement leftColums = driver.findElement(By.cssSelector("div[class='leftColumnContent']"));
			String leftSide = leftColums.getText();
			
//			List<WebElement> specimentE = driver.findElements(By.tagName("p"));
//			if(!specimentE.isEmpty()) {
//				speciment=specimentE.get(0).getText();
//			}
//			
//			List<WebElement> containerE = driver.findElements(By.xpath("//div/div[2]"));
//			if(!containerE.isEmpty()) {
//				container=containerE.get(0).getText();
//			}
			datalist.add(new String[] {name,component, leftSide,company, city, country, link});

		} catch (Exception e) {
			System.out.println("EEEEEEE+ " + rowIndex);
			return false;
		}
		return true;

	}

	public void clickAction(WebElement element) {
		JavascriptExecutor js = (JavascriptExecutor) driver;
		js.executeScript("arguments[0].click();", element);
	}

	public void selectComboValue(final String elementName, final String value) {
		final Select selectBox = new Select(driver.findElement(By.cssSelector(elementName)));
		selectBox.selectByValue(value);
	}

	private void appendExcelFile(String fileName, String sheetName) {
		Workbook workbook = null;
		Sheet sheet;
		try {
			File file = new File("D:\\Freelancer\\thoan_excel\\Results\\" + fileName);
			if (!file.exists()) {
				workbook = new XSSFWorkbook(); // new HSSFWorkbook() for generating `.xls` file
			} else {
				FileInputStream fip = new FileInputStream(file);
				workbook = new XSSFWorkbook(fip);
			}

			if (workbook.getNumberOfSheets() == 0) {
				sheet = workbook.createSheet("Results");
			} else {
				sheet = workbook.getSheetAt(0);
			}

			int rowNum = sheet.getPhysicalNumberOfRows();
			if (!sheetName.isEmpty()) {
				sheet.createRow(rowNum++).createCell(0).setCellValue(sheetName);
			}
			for (String[] d : datalist) {
				Row row = sheet.createRow(rowNum++);
				for (int i = 0; i < d.length; i++) {
					row.createCell(i).setCellValue(d[i]);
				}
			}

			FileOutputStream fileOut = null;
			try {
				fileOut = new FileOutputStream(file);
				workbook.write(fileOut);
				datalist.clear();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} finally {
				if (fileOut != null) {
					try {
						fileOut.close();
					} catch (IOException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
				}
			}

		} catch (IOException e) {
			e.printStackTrace();
		}

		// Create Other rows and cells with employees data

		// Closing the workbook
		try {
			workbook.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	private void writeExcelFile(String fileName, String sheetName) {

		Workbook workbook = new XSSFWorkbook(); // new HSSFWorkbook() for generating `.xls` file
		Sheet sheet = workbook.createSheet("Results");

		int rowNum = 0;
		if (!sheetName.isEmpty()) {
			sheet.createRow(rowNum++).createCell(0).setCellValue(sheetName);
		}
		for (String[] d : datalist) {
			Row row = sheet.createRow(rowNum++);
			for (int i = 0; i < d.length; i++) {
				row.createCell(i).setCellValue(d[i]);
			}
		}

		FileOutputStream fileOut = null;
		try {
			fileOut = new FileOutputStream("D:\\Freelancer\\thoan_excel\\Results\\" + fileName + ".xlsx");
			workbook.write(fileOut);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} finally {
			if (fileOut != null) {
				try {
					fileOut.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		}

		// Closing the workbook
		try {
			workbook.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}
