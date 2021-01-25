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

public class Dynacare_ca {
	static WebDriver driver;
	List<String[]> datalist = new ArrayList<>();
	String fileName = "dynacare.ca.xlsx";
	String fileResult = "dynacare.ca_result.xlsx";

	public Dynacare_ca() {
		System.setProperty("webdriver.chrome.driver", "D:\\Freelancer\\thoan_excel\\chromedriver\\chromedriver.exe");
		driver = new ChromeDriver();
		driver.manage().window().maximize();
	}

	public static void main(String[] args) {

		String url = "https://www.dynacare.ca/find-a-location.aspx";
		Dynacare_ca t = new Dynacare_ca();
//		 t.processWeb(url);
		t.processExcel();

	}

	public void processWeb(String url) {
		driver.get(url);
		try {
			Thread.sleep(3000);
			List<WebElement> items = driver.findElements(By.cssSelector("div[class='viewDetails']"));
			for (int i=0; i<items.size();i++) {
				String link = items.get(i).findElement(By.tagName("a")).getAttribute("href");
				datalist.add(new String[] { link});
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

				for (int i = 47; i < table1Sheet.getPhysicalNumberOfRows(); i++) {
					try {

						Cell cell = table1Sheet.getRow(i).getCell(0);
						if (cell == null) {
							continue;
						}

						String url = table1Sheet.getRow(i).getCell(0).getStringCellValue();
						
						boolean isNext = processPage(url, i + 1);
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

	public boolean processPage(String url, int rowIndex) {

		try {
			driver.get(url);
			Thread.sleep(2000);
			
			String name="";
			String address="";
			String city="";
			String postcode="";
			String phone="";
			String fax="";
			String operationalHours="";
			String emergencyOperationalHours="";
			
			List<WebElement> nameE = driver.findElements(By.id("p_lt_ctl06_pageplaceholder_p_lt_ctl01_DYN_FindALocationDetails_Enqueue_STN"));
			if(!nameE.isEmpty()) {
				name=nameE.get(0).getText();
			}
			
			List<WebElement> addressE = driver.findElements(By.id("p_lt_ctl06_pageplaceholder_p_lt_ctl01_DYN_FindALocationDetails_Enqueue_Address"));
			if(!addressE.isEmpty()) {
				address=addressE.get(0).getText();
			}
			
			List<WebElement> cityE = driver.findElements(By.id("p_lt_ctl06_pageplaceholder_p_lt_ctl01_DYN_FindALocationDetails_Enqueue_CityAndProvince"));
			if(!cityE.isEmpty()) {
				city=cityE.get(0).getText();
			}
			
			List<WebElement> postcodeE = driver.findElements(By.id("p_lt_ctl06_pageplaceholder_p_lt_ctl01_DYN_FindALocationDetails_Enqueue_PostalCode"));
			if(!postcodeE.isEmpty()) {
				postcode=postcodeE.get(0).getText();
			}
			
			List<WebElement> phoneE = driver.findElements(By.id("p_lt_ctl06_pageplaceholder_p_lt_ctl01_DYN_FindALocationDetails_Enqueue_Phone"));
			if(!phoneE.isEmpty()) {
				phone=phoneE.get(0).getText();
			}
			
			List<WebElement> faxE = driver.findElements(By.id("p_lt_ctl06_pageplaceholder_p_lt_ctl01_DYN_FindALocationDetails_Enqueue_Fax"));
			if(!faxE.isEmpty()) {
				fax=faxE.get(0).getText();
			}
			
			List<WebElement> operationalHoursE = driver.findElements(By.id("operationalHours"));
			if(!operationalHoursE.isEmpty()) {
				operationalHours=operationalHoursE.get(0).getText();
			}
			
			List<WebElement> emergencyOperationalHoursE = driver.findElements(By.id("emergencyOperationalHours"));
			if(!emergencyOperationalHoursE.isEmpty()) {
				emergencyOperationalHours=emergencyOperationalHoursE.get(0).getText();
			}
			
			String all = driver.findElement(By.id("ldMainContent")).getText();
			
			datalist.add(new String[] {name,address, city,postcode, phone, fax, operationalHours, emergencyOperationalHours, all});

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
