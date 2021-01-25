package selenium;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.Date;
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
import org.openqa.selenium.PageLoadStrategy;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.Select;

public class Lawyersfirms {
	static WebDriver driver;
	List<String[]> datalist = new ArrayList<>();
	String fileName = "lawyersfirms.xlsx";
	String fileResult = "lawyersfirms_result.xlsx";

	public Lawyersfirms() {
		System.setProperty("webdriver.chrome.driver", "D:\\Freelancer\\thoan_excel\\driver87\\chromedriver.exe");
		driver = new ChromeDriver();
		driver.manage().window().maximize();
	}

	public static void main(String[] args) {

		String url = "https://www.lawyersfirms.com.au/law-types/family-law";
		Lawyersfirms t = new Lawyersfirms();
		t.processWeb(url);
		// t.processExcel();

	}

	public void processWeb(String url) {
		driver.get(url);
	
			try {
				Thread.sleep(3000);
			} catch (InterruptedException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}
			
			List<String> urls = new ArrayList<>();
			urls.add(url);
			List<WebElement> nextPageE = driver
					.findElements(By.xpath("//*[@id=\"main-wrapper\"]/section[1]/div/div[2]/a"));
			for (WebElement item : nextPageE) {
				String nextUrl = item.getAttribute("href");
				urls.add(nextUrl);
			}
			for(String nextUrl : urls) {
				try {
					if (processPage(nextUrl)) {
						System.out.println("Write page:"+ nextUrl + " size: " + datalist.size());
						appendExcelFile(fileName, "");
					}
				} catch (Exception e) {
					e.printStackTrace();
				}
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

				for (int i = 0; i < table1Sheet.getPhysicalNumberOfRows(); i++) {
					try {

						Cell cell = table1Sheet.getRow(i).getCell(0);
						if (cell == null) {
							continue;
						}

						String url = table1Sheet.getRow(i).getCell(0).getStringCellValue();

						// boolean isNext = processPage(url, i + 1);
						if (datalist.size() == 100) {
							appendExcelFile(fileResult, "");
							Thread.sleep(3000);
							System.out.println("==========" + (i + 1));
						}
						// if(!isNext) {
						// break;
						// }

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

	public boolean processPage(String nextUrl) {
		try {
			driver.get(nextUrl);
			Thread.sleep(3000);
			List<WebElement> items = driver.findElements(By.xpath("//*[@id=\"main-wrapper\"]/section[1]/div/div[1]/div"));
			for (WebElement item : items) {
				try {
					WebElement contentE = item.findElement(By.cssSelector("div[class='proerty_content']"));

					WebElement nameE = contentE.findElement(By.cssSelector("div[class='proerty_text']"));
					String name = nameE.getText();
					String link = nameE.findElement(By.tagName("a")).getAttribute("href");
					WebElement addressE = contentE.findElement(By.cssSelector("p[class='property_add']"));
					String address = addressE.getText();
					WebElement infosE = contentE.findElement(By.cssSelector("div[class='list-fx-features']"));
					String infos = infosE.getText();
					datalist.add(new String[] { name, link, address, infos });
				} catch (Exception e1) {
					e1.printStackTrace();
				}
			}
		}catch (Exception e) {
			
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
