package selenium;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

import javax.crypto.Mac;

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
import org.openqa.selenium.Keys;
import org.openqa.selenium.PageLoadStrategy;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.Select;

public class Bidclerk_Com {
	static WebDriver driver;
	List<String[]> datalist = new ArrayList<>();
	List<Map<String, String[]>> datalist2 = new ArrayList<>();
	String fileName = "bidclerk.com.xlsx";
	String fileResult = "bidclerk.com_result5.xlsx";

	public Bidclerk_Com() {
		System.setProperty("webdriver.chrome.driver", "D:\\Freelancer\\thoan_excel\\driver79\\chromedriver.exe");

		ChromeOptions options = new ChromeOptions();
		options.addArguments("enable-automation");
		options.addArguments("--headless");
		options.addArguments("--window-size=1920,1080");
		options.addArguments("--no-sandbox");
		options.addArguments("--disable-extensions");
		options.addArguments("--dns-prefetch-disable");
		options.addArguments("--disable-gpu");
		options.setPageLoadStrategy(PageLoadStrategy.NORMAL);
		driver = new ChromeDriver(options);
		driver.manage().window().maximize();
	}

	public static void main(String[] args) {

		String url = "https://www.bidclerk.com/advanced-search-project/-999.html";
		Bidclerk_Com t = new Bidclerk_Com();
		// t.processWeb(url);
		t.processExcel();

	}

	public void doLogin() {
		try {
			String loginUrl = "https://www.bidclerk.com/login.html";
			driver.get(loginUrl);
			Thread.sleep(3000);
			driver.findElement(By.xpath("//*[@id=\"login-form-container\"]/form/input[1]"))
					.sendKeys("dblatt@capstackpartners.com");
			driver.findElement(By.xpath("//*[@id=\"login-form-container\"]/form/input[2]")).sendKeys("wqsaxz78");
			driver.findElement(By.id("home-login-submit-id")).click();
			Thread.sleep(3000);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void processWeb(String url) {
		doLogin();

		try {
			driver.get(url);
			Thread.sleep(3000);
			WebElement removeFilter = driver.findElement(By.xpath("//*[@id=\"bc-trail\"]/span[5]/a"));
			clickAction(removeFilter);
			Thread.sleep(5000);
			String activePage = "";
			String[] matching = new String[] { "1190" };
			while (true) {
				WebElement activePageE = driver.findElement(By.cssSelector("a[class='paginate_active']"));
				activePage = activePageE.getText().replace(",", "");

				int currentPageNumber = Integer.parseInt(activePage);

				if (currentPageNumber == 1190) {
					processPage();
					break;

				}

				List<WebElement> nextBs = driver.findElements(By.xpath("//*[@id=\"search-results_next\"]"));
				if (nextBs.isEmpty()) {
					break;
				} else {
					clickAction(nextBs.get(0));
				}
				Thread.sleep(5000);
			}

		} catch (Exception e) {
			e.printStackTrace();
		}

		System.out.println("===============Finish====================");
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

				for (int i = 86550; i < table1Sheet.getPhysicalNumberOfRows(); i++) {
					try {

						Cell cell = table1Sheet.getRow(i).getCell(0);
						if (cell == null) {
							continue;
						}

						String link = table1Sheet.getRow(i).getCell(0).getStringCellValue();

						if(!processWeb2(link, i + 1)) {
							break;
						}

					} catch (Exception e) {
						System.out.println("EEEEEERRRRRRRRRR: " + (i + 1));
						e.printStackTrace();
					}
				}
				System.out.println("===============Error====================");
				for (String error : errorList) {
					System.out.println(error);
				}
				
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		
		appendExcelFile(fileResult, "");
		System.out.println("===============Finish====================");

	}

	List<String> errorList = new ArrayList<>();

	public boolean processWeb2(String link, int rowIndex) {

		try {
			driver.get(link);
			Thread.sleep(1000);
			List<WebElement> items = driver.findElement(By.cssSelector("div[class='table_wrapper']"))
					.findElements(By.tagName("tr"));
			items.remove(0);
			
			boolean isOwner = false;
			List<String> contacts = new ArrayList<>();
			for (WebElement item : items) {

				List<WebElement> tds = item.findElements(By.tagName("td"));
				String role = tds.get(2).getText();
				if (!role.toUpperCase().contains("OWNER")) {
					continue;
				} else {
					isOwner = true;
					String name = tds.get(0).getText();
					String company = tds.get(1).getText();

					String phone = tds.get(3).getText();
					String email = tds.get(5).getText();
					contacts.add(name);
					contacts.add(company);
					contacts.add(role);
					contacts.add(phone);
					contacts.add(email);
				}
			}
			
			if (!isOwner) {
				Map<String, String[]> data = new HashMap<>();
				data.put("rowIndex", new String[] { "" + rowIndex });
				data.put("link", new String[] { link });
				data.put("projectId", new String[] {});
				data.put("projectName", new String[] {});
				data.put("location1", new String[] {});
				data.put("description", new String[] {});
				data.put("startDate", new String[] {});
				data.put("buidingUse", new String[] {});
				data.put("estimatedValue", new String[] {});
				data.put("projectContacts", new String[] {});

				datalist2.add(data);
			} else {
				String[] projectContacts = new String[contacts.size()];
				for(int i=0;i<contacts.size();i++) {
					projectContacts[i]=contacts.get(i);
				}
				
				String projectId = driver.findElement(By.id("project-id")).getAttribute("data-entity-id");
				String projectName = driver.findElement(By.id("project-title")).getAttribute("title");
				String estimatedValue = "";
				String buildingUse = "";
				String description = driver.findElement(By.xpath("//*[@id=\"projectDescriptionJoyride\"]/div"))
						.getText();
				List<WebElement> projectDetailsE = driver.findElement(By.cssSelector("ul[class='projectDetails']"))
						.findElements(By.tagName("li"));
				for (WebElement item : projectDetailsE) {
					String value = item.getText();
					if (item.getText().contains("Use")) {
						buildingUse = value.replace("Building Use", "");
						if (buildingUse.length() > 1) {
							buildingUse = buildingUse.substring(1);
						}
						continue;
					} else if (item.getText().contains("Estimated")) {
						estimatedValue = value.replace("Estimated Value", "");
						if (estimatedValue.length() > 1) {
							estimatedValue = estimatedValue.substring(1);
						}
					}
				}
				List<WebElement> startDateE = driver.findElements(By.xpath("//*[@id=\"detail-events\"]/tbody/tr"));
				String startDate = "";
				for (WebElement item : startDateE) {
					if (item.getText().contains("Start Date")) {
						startDate = item.getText().replace("Est.Start Date", "");
						if (startDate.length() > 1) {
							startDate = startDate.substring(1);
						}
						break;
					}
				}
				WebElement locationE = driver.findElement(By.cssSelector("div[class='locationInfo']"));
				String location1 = locationE.findElements(By.tagName("h4")).get(1).getText();
				// String location2=locationE.findElement(By.tagName("p")).getText();

				Map<String, String[]> data = new HashMap<>();
				data.put("rowIndex", new String[] { "" + rowIndex });
				data.put("link", new String[] { link });
				data.put("projectId", new String[] { projectId });
				data.put("projectName", new String[] { projectName });
				data.put("location1", new String[] { location1 });
				// data.put("location2", new String[] {location2});
				data.put("description", new String[] { description });
				data.put("startDate", new String[] { startDate });
				data.put("buidingUse", new String[] { buildingUse });
				data.put("estimatedValue", new String[] { estimatedValue });
				data.put("projectContacts", projectContacts);

				datalist2.add(data);
			}
			if (rowIndex % 50 == 0) {
				appendExcelFile2(fileResult, "");
				System.out.println("===============Write row:" + (rowIndex));
			}

		} catch (Exception e) {
			e.printStackTrace();
			System.out.println("ERRRRRR: " + link);
			errorList.add("" + rowIndex);
			try {
				Thread.sleep(5000);
			}catch (Exception ex) {
				ex.printStackTrace();
			}
			if(errorList.size()==20) {
				return false;
			}
		}
		
		return true;

	}

	public boolean processPage() {
  
		try {
			WebElement activePageE = driver.findElement(By.cssSelector("a[class='paginate_active']"));
			String activePage = activePageE.getText();

			List<WebElement> items = driver.findElements(By.xpath("//*[@id=\"search-results\"]/tbody/tr"));
			for (WebElement item : items) {
				String link = item.findElement(By.tagName("a")).getAttribute("href");
				datalist.add(new String[] { link, activePage });
			}
			System.out.println("=======Write: " + activePage);
			appendExcelFile(fileName, "");

		} catch (Exception e) {
			e.printStackTrace();
			try {
				Thread.sleep(5000);
				datalist.clear();
				WebElement activePageE = driver.findElement(By.cssSelector("a[class='paginate_active']"));
				String activePage = activePageE.getText();
				List<WebElement> items = driver.findElements(By.xpath("//*[@id=\"search-results\"]/tbody/tr"));
				for (WebElement item : items) {
					String link = item.findElement(By.tagName("a")).getAttribute("href");
					datalist.add(new String[] { link, activePage });
				}
				System.out.println("=======Write: " + activePage);
				appendExcelFile(fileName, "");
			} catch (InterruptedException e1) {

				e1.printStackTrace();
				return false;
			}
			return true;
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

	private void appendExcelFile2(String fileName, String sheetName) {
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
			int cellNum = 0;
			if (!sheetName.isEmpty()) {
				sheet.createRow(rowNum++).createCell(cellNum).setCellValue(sheetName);
			}

			for (Map<String, String[]> d : datalist2) {
				Row row = sheet.createRow(rowNum++);
				String link = d.get("link")[0];
				writeExcelRow(row, cellNum, d.get("rowIndex"), link);
				cellNum++;
				writeExcelRow(row, cellNum, d.get("link"), link);
				cellNum++;
				writeExcelRow(row, cellNum, d.get("projectId"), link);
				cellNum++;
				writeExcelRow(row, cellNum, d.get("projectName"), link);
				cellNum++;
				writeExcelRow(row, cellNum, d.get("location1"), link);

				cellNum++;
				writeExcelRow(row, cellNum, d.get("description"), link);
				cellNum++;
				writeExcelRow(row, cellNum, d.get("startDate"), link);
				cellNum++;
				writeExcelRow(row, cellNum, d.get("buidingUse"), link);
				cellNum++;
				writeExcelRow(row, cellNum, d.get("estimatedValue"), link);
				cellNum++;
				writeExcelRow(row, cellNum, d.get("projectContacts"), link);
				cellNum = 0;
			}

			FileOutputStream fileOut = null;
			try {
				fileOut = new FileOutputStream(file);
				workbook.write(fileOut);
				datalist2.clear();
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

	private void writeExcelRow(Row row, int cellNum, String[] data, String link) {
		if (data.length > 100) {
			System.out.println("Over limit:" + link);
		}
		for (String d : data) {
			row.createCell(cellNum++).setCellValue(d);
		}
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

}
