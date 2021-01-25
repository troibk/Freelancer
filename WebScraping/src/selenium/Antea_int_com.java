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

public class Antea_int_com {
	static WebDriver driver;
	List<String[]> datalist = new ArrayList<>();
	String fileName = "antea-int.com.xlsx";
	String fileResult = "antea-int.com_result.xlsx";

	public Antea_int_com() {
		System.setProperty("webdriver.chrome.driver", "D:\\Freelancer\\thoan_excel\\chromedriver\\chromedriver.exe");
		driver = new ChromeDriver();
		driver.manage().window().maximize();
	}

	public static void main(String[] args) {

		String url = "https://antea-int.com/find-a-member/";
		Antea_int_com t = new Antea_int_com();
		t.processWeb(url);
		// t.processExcel();

	}

	public void processWeb(String url) {
		driver.get(url);
		try {
			Thread.sleep(3000);
			WebElement es = driver.findElement(By.xpath("//*[@id=\"soflow\"]"));
			List<WebElement> countries = es.findElements(By.tagName("option"));
			List<String> countriesT= new ArrayList<>();
			for(int i=1;i<countries.size();i++) {
				String country = countries.get(i).getAttribute("value");
				countriesT.add(country);
			}
			for (String country: countriesT) {
				driver.get(url);
				Thread.sleep(3000);
				try {
					
					selectComboValue(driver.findElement(By.xpath("//*[@id=\"soflow\"]")), country);
					Thread.sleep(1000);
					WebElement cityE = driver.findElements(By.xpath("//*[@id=\"soflow\"]")).get(1);
					List<WebElement> cities = cityE.findElements(By.tagName("option"));
					for (int k = 1; k < cities.size(); k++) {
						String city= cities.get(k).getAttribute("value");
						selectComboValue(cityE, city);
						Thread.sleep(2000);
						List<WebElement> points = driver.findElements(By.xpath(
								"//*[@id=\"mapa-miembros-antea-content\"]/div[2]/div/div/div/div[1]/div[3]/div/div[3]/div"));
						for (WebElement point : points) {
							List<WebElement> imgs = point.findElements(By.tagName("img"));
							if (!imgs.isEmpty()) {
								clickAction(imgs.get(0));
								Thread.sleep(1000);
							}
						}
						Thread.sleep(1000);
						List<WebElement> data = driver.findElements(By.xpath(
								"//*[@id=\"mapa-miembros-antea-content\"]/div[2]/div/div/div/div[1]/div[3]/div/div[4]/div"));
						for (int j = 0; j < data.size(); j++) {
							String v = data.get(j).getText();
							if (!v.isEmpty()) {
								datalist.add(new String[] { v, country, city });
							}
						}
						
					}
					appendExcelFile(fileName, "");
					System.out.println("xxxxxxxx " + country);
				} catch (Exception e2) {
					e2.printStackTrace();
					System.out.println("===========" + country);
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		}

		System.out.println("===============Finish====================");
	}

	public void clickAction(WebElement element) {
		JavascriptExecutor js = (JavascriptExecutor) driver;
		js.executeScript("arguments[0].click();", element);
	}

	public void selectComboValue(final WebElement element, final String value) {
		final Select selectBox = new Select(element);
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
