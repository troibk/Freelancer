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

public class Realestate_com_au {
	static WebDriver driver;
	List<String[]> datalist = new ArrayList<>();
	String fileName = "realestate.com.au.xlsx";
	String fileResult = "realestate.com.au.xlsx";

	public Realestate_com_au() {
		System.setProperty("webdriver.chrome.driver", "D:\\Freelancer\\thoan_excel\\driver81\\chromedriver.exe");
		ChromeOptions opt = new ChromeOptions();
		opt.setBinary("C:\\Program Files (x86)\\Google\\Chrome Beta\\Application\\chrome.exe");
		 driver = new ChromeDriver(opt);
		driver.manage().window().maximize();
		driver.manage().deleteAllCookies();
	}

	public static void main(String[] args) {

		String url = "https://www.realestate.com.au/sold/";
		Realestate_com_au t = new Realestate_com_au();
		t.processWeb(url);
		// t.processExcel();

	}

	public void processWeb(String url) {
		driver.get(url);
		try {
			Thread.sleep(3000);
			WebElement searchInput = driver.findElement(By.xpath("//*[@id=\"where\"]"));
			searchInput.sendKeys("3088");
			Thread.sleep(500);
			WebElement searchBtn = driver
					.findElement(By.xpath("/html/body/div[1]/div[1]/div[1]/form/div/div[1]/div/div/button"));
			clickAction(searchBtn);
			int pageNumber = 0;
			while (true) {

				try {
					pageNumber++;
					Thread.sleep(5000);
					if (processPage()) {
						appendExcelFile(fileName, "");
						System.out.println("Write page:" + pageNumber);
						List<WebElement> nextPageE = driver
								.findElements(By.cssSelector("div[class='pagination__link-next-wrapper ']"));
						if (!nextPageE.isEmpty()) {
							clickAction(nextPageE.get(0).findElement(By.tagName("a")));
						}
					} else {
						break;
					}
				} catch (Exception e) {
					e.printStackTrace();
					break;
				}

			}

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

	public boolean processPage() {
		String soldDate = null;
		List<WebElement> items = driver.findElements(By.cssSelector("div[class='residential-card__content-wrapper']"));
		for (WebElement item : items) {
			WebElement addressE = item
					.findElement(By.cssSelector("div[class='details-link residential-card__details-link']"));
			String address = addressE.getText();
			WebElement soldDateE = item.findElement(By.cssSelector("div[class='piped-content']"));
			soldDate = soldDateE.getText().replace("Sold on ", "").trim();
			WebElement typeE = driver.findElement(By.cssSelector("span[class='residential-card__property-type']"));
			String type = typeE.getText();
			datalist.add(new String[] { soldDate, address, type });
		}

		if (soldDate != null) {

			SimpleDateFormat formatter = new SimpleDateFormat("dd MMM yyyy");

			try {
				Date soldDateO = formatter.parse(soldDate);
				String date60S = LocalDate.now().minusDays(60).toString();
				Date date60 = formatter.parse(date60S);
				if (soldDateO.compareTo(date60) <= 0) {
					return false;
				}

			} catch (ParseException e) {
				e.printStackTrace();
				return false;
			}
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
