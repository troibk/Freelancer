package upwork.haworker25;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedCondition;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.openqa.selenium.JavascriptExecutor;

public class Scraping {
	static WebDriver driver;
	String url;
	String folderPath;
	String resultName = "results";
	private static List<String[]> datalist = new ArrayList<>();

	public Scraping(String folderPath, String url) {
		this.url = url;
		this.folderPath = folderPath;
		String projectPath = System.getProperty("user.dir");
		System.setProperty("webdriver.chrome.driver", projectPath + "\\resources\\chromedriver.exe");
		driver = new ChromeDriver();
		driver.manage().window().maximize();
	}

	public void doLogin() {
		try {

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void processWeb() {
		UserGUI.setLog("Connecting....");
		try {
		} catch (Exception e) {
			e.printStackTrace();
		}
		System.out.println("===============Finish====================");
	}

	public int processPage() {
		List<WebElement> rows = driver.findElements(By.xpath(""));
		for (int i = 1; i < rows.size(); i++) {
			List<WebElement> tds = rows.get(i).findElements(By.tagName("td"));
			String name = tds.get(1).getText();
			String link = tds.get(1).findElement(By.tagName("a")).getAttribute("href");
			String firm = tds.get(2).getText();
			String country = tds.get(3).getText();
			String[] data = new String[4];
			data[0] = link;
			data[1] = name;
			data[2] = firm;
			data[3] = country;
			datalist.add(data);
		}
		return rows.size();
	}

	public void processExcel() {
		try {

		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	public String processWeb(String link, int rowIndex) {
		String result = "";

		try {

		} catch (Exception e) {
			e.printStackTrace();
			System.out.println("ERRRRRR: " + rowIndex);
			return null;
		}
		return result;
	}
	
	public void clickAction(WebElement element) {
		JavascriptExecutor js = (JavascriptExecutor)driver;
		js.executeScript("arguments[0].click();", element);
	}

	public void waitForPageLoad() {
		ExpectedCondition<Boolean> pageLoadCondition = new ExpectedCondition<Boolean>() {
			public Boolean apply(WebDriver driver) {
				return ((JavascriptExecutor) driver).executeScript("return document.readyState").equals("complete");
			}
		};
		WebDriverWait wait = new WebDriverWait(driver, 30);
		wait.until(pageLoadCondition);
	}

	public WebElement waitForElementVisible(int type, String value) {
		WebDriverWait wait = new WebDriverWait(driver, 10);
		WebElement element = null;
		switch (type) {
		case 1:
			element = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id(value)));
			break;
		case 2:
			element = wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector(value)));
			break;
		case 3:
			element = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(value)));
			break;
		default:
			break;
		}

		return element;
	}

	public void selectComboValue(final String elementName, final String value) {
		final Select selectBox = new Select(driver.findElement(By.cssSelector(elementName)));
		selectBox.selectByValue(value);
	}

	private void appendExcelFile(String sheetName) {
		Workbook workbook = null;
		Sheet sheet;
		try {
			File file = new File(folderPath + "\\" + resultName);
			if (!file.exists()) {
				workbook = new XSSFWorkbook();
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
				e.printStackTrace();
			} finally {
				if (fileOut != null) {
					try {
						fileOut.close();
					} catch (IOException e) {
						e.printStackTrace();
					}
				}
			}

		} catch (IOException e) {
			e.printStackTrace();
		}

		try {
			workbook.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	private void writeExcelFile(String sheetName) {

		Workbook workbook = new XSSFWorkbook();
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
			File file = new File(folderPath + "\\" + resultName);
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
