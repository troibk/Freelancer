package selenium;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
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
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

public class Angel_co {
	static WebDriver driver;
	List<String[]> datalist = new ArrayList<>();
	String fileName = "angel.co.xlsx";
	String fileResult = "angel.co_result.xlsx";

	public Angel_co() {
		System.setProperty("webdriver.chrome.driver", "D:\\Freelancer\\thoan_excel\\chromedriver\\chromedriver.exe");
		driver = new ChromeDriver();
		driver.manage().window().maximize();
	}

	public static void main(String[] args) {

		String url = "https://angel.co/jobs#find/f!%7B%22roles%22%3A%5B%22Software%20Engineer%22%2C%22Hardware%20Engineer%22%2C%22Mechanical%20Engineer%22%2C%22Systems%20Engineer%22%2C%22Data%20Scientist%22%5D%2C%22locations%22%3A%5B%221688-United%20States%22%5D%2C%22types%22%3A%5B%22contract%22%5D%7D";
		Angel_co t = new Angel_co();
		t.processWeb(url);
		// t.processExcel();

	}

	public void doLogin() {
		try {
			String loginUrl = "https://angel.co/login?after_sign_in=https%3A%2F%2Fangel.co%2Fonboarding%2Fstart_job_profile";
			driver.get(loginUrl);
			Thread.sleep(3000);
			driver.findElement(By.id("user_email")).sendKeys("troibk@gmail.com");
			driver.findElement(By.id("user_password")).sendKeys("Tlvctg2@");

			WebElement btn = driver.findElement(By.xpath("//*[@id=\"new_user\"]/div[2]/input"));
			JavascriptExecutor js = (JavascriptExecutor) driver;
			js.executeScript("arguments[0].click();", btn);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void processWeb(String url) {

		doLogin();
		try {
			driver.get(url);
			Thread.sleep(3000);
			JavascriptExecutor js = (JavascriptExecutor) driver;
			while (true) {
				js.executeScript("window.scrollBy(0,1000)");
				Thread.sleep(3000);
				List<WebElement> es = driver.findElements(By.cssSelector("div[class='header-info']"));
			
				if (es.size()>=527) {
					for (int i =0;i<es.size();i++) {
						if(i>0) {
							js.executeScript("arguments[0].click();", es.get(i));
							Thread.sleep(1000);
						}
						WebElement info = es.get(i).findElement(By.cssSelector("a[class='startup-link']"));
						String link = info.getAttribute("href");
						String name = info.getText();
						List<WebElement> moreDetails = driver.findElements(By.cssSelector("div[class='startup-info-table']"));
						List<WebElement> activeE = moreDetails.get(i).findElements(By.cssSelector("div[class='active']"));
						String active="";
						if(!activeE.isEmpty()) {
							active= activeE.get(0).getText();
						}
						String[] data=new String[3];
						data[0]=link;
						data[1]=active;
						data[2]=name;
						datalist.add(data);
					}
					break;
				}
			}

		} catch (Exception e) {
			e.printStackTrace();
		}
		appendExcelFile(fileName, "");
		System.out.println("===============Finish====================");
	}

	public int processPage() {
		List<WebElement> rows = driver.findElements(By.xpath("//*[@id=\"mainForm\"]/div[3]/div/table/tbody/tr"));
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
		doLogin();
		try {
			FileInputStream inputFile = new FileInputStream(
					new File("D:\\Freelancer\\thoan_excel\\Results\\" + fileName));
			XSSFWorkbook wb = new XSSFWorkbook(inputFile);
			XSSFSheet table1Sheet = wb.getSheetAt(0);

			if (table1Sheet == null) {
				System.out.println("KO CO SHEET :" + table1Sheet);
			} else {

				for (int i = 209; i < table1Sheet.getPhysicalNumberOfRows(); i++) {
					try {

						Cell cell = table1Sheet.getRow(i).getCell(0);
						if (cell == null) {
							continue;
						}

						String link = table1Sheet.getRow(i).getCell(0).getStringCellValue();
						String name = table1Sheet.getRow(i).getCell(1).getStringCellValue();
						String firm = table1Sheet.getRow(i).getCell(2).getStringCellValue();
						String country = table1Sheet.getRow(i).getCell(3).getStringCellValue();
						String email = processWeb(link, i + 1);
						if (email == null) {
							break;
						}
						String[] data = new String[4];
						data[0] = firm;
						data[1] = name;
						data[2] = country;
						data[3] = email;
						datalist.add(data);
						if ((i + 1) % 50 == 0) {
							appendExcelFile(fileResult, "");
							System.out.println("===============Write row:" + (i + 1));
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

	public String processWeb(String link, int rowIndex) {
		String result = "";

		try {
			driver.get(link);
			Thread.sleep(2000);
			List<WebElement> emaislE = driver.findElements(By.id("ctl00_MainContent_lnkEmailAddress"));
			if (!emaislE.isEmpty()) {
				result = emaislE.get(0).getText();
			}
		} catch (Exception e) {
			e.printStackTrace();
			System.out.println("ERRRRRR: " + rowIndex);
			return null;
		}
		return result;
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

	private void appendExcelFile(String fileName, String sheetName) {
		Workbook workbook = null;
		Sheet sheet;
		try {
			File file = new File("D:\\Freelancer\\thoan_excel\\Results\\" + fileName);
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
