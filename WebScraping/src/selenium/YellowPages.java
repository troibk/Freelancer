package selenium;

import java.awt.Button;
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

public class YellowPages {
	static WebDriver driver;
	List<String[]> datalist = new ArrayList<>();
	String fileName = "accountants_vic.xlsx";
	String fileResult = "int-bar.org_result.xlsx";

	public YellowPages() {
		System.setProperty("webdriver.chrome.driver", "D:\\Freelancer\\thoan_excel\\chromedriver\\chromedriver.exe");
		driver = new ChromeDriver();
		driver.manage().window().maximize();
	}

	public static void main(String[] args) {

		String url = "https://www.yellowpages.com.au/search/listings?clue=Accountants+%26+auditors&locationClue=vic&lat=&lon=&selectedViewMode=list";
		YellowPages t = new YellowPages();
		t.processWeb(url);
		// t.processExcel();

	}

	public void doLogin() {
		try {
			String loginUrl = "https://www.ibanet.org/Access/SignIn.aspx?url=/MySite/";
			driver.get(loginUrl);
			Thread.sleep(3000);
			driver.findElement(By.id("ctl00_MainContent_tbxUserName")).sendKeys("1404140");
			driver.findElement(By.id("ctl00_MainContent_tbxPassword")).sendKeys("IRglobal20!9");
			driver.findElement(By.id("ctl00_MainContent_btnSignIn")).click();
			Thread.sleep(3000);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void doCheckboxCaptcha(String url) {
		try {
			driver.get(url);
			Thread.sleep(3000);
			WebElement iframe = driver.findElement(By.xpath("/html/body/div[2]/div/div/div[3]/form/div[1]/div/div/iframe"));
			driver.switchTo().frame(iframe);
			WebElement checkbox = driver.findElement(By.xpath("//*[@id=\"recaptcha-anchor\"]"));
			checkbox.click();
			WebElement submit = driver.findElement(By.cssSelector("button[class='submit']"));
			submit.click();
			Thread.sleep(3000);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void processWeb(String url) {

		doCheckboxCaptcha(url);
		try {
			int currentPage=0;
			while (true) {
				currentPage++;
				int pageSize = processPage(currentPage);
				if (pageSize > 0) {
					
					System.out.println("===========Current Page: " + currentPage + ", Size: " + datalist.size());
					appendExcelFile(fileName, "");
					Thread.sleep(5000);
					List<WebElement> pagingBtn = driver
							.findElement(By.cssSelector("div[class='search-pagination-container']")).findElements(By.tagName("a"));
					boolean isNext=false;
					WebElement nextBtn = pagingBtn.get(pagingBtn.size()-1);
					if (nextBtn.getText().equals("Next")) {
							JavascriptExecutor js = ((JavascriptExecutor) driver);
							js.executeScript("arguments[0].click();", nextBtn);
							isNext = true;
							break;
						}
					
					if (!isNext) {
						break;
					}
					Thread.sleep(5000);
				} else {
					break;
				}
			}

		} catch (Exception e) {
			e.printStackTrace();
		}
		System.out.println("===============Finish====================");
	}

	public int processPage(int currentPage) {
		List<WebElement> rows = driver.findElements(By.xpath("//*[@id=\"search-results-page\"]/div[1]/div/div[4]/div/div/div[2]/div/div[2]/div[2]/div/div"));
		for (int i = 2; i < rows.size(); i++) {
			String name = rows.get(i).findElement(By.cssSelector("a[class='listing-name']")).getText();
			List<WebElement> addressE = rows.get(i).findElements(By.cssSelector("div[class='poi-and-body']"));
			String address ="";
			if(!addressE.isEmpty()) {
				address = addressE.get(0).getText();
			}
			List<WebElement> actions = rows.get(i).findElements(By.cssSelector("div[class='call-to-action   ']"));
			String phone ="";
			String email ="";
			String web ="";
			for(WebElement a : actions) {
			
			String href = a.findElement(By.tagName("a")).getAttribute("href");

			if(href.contains("tel")) {
				phone=href;
			}else if(href.contains("mailto")) {
				email =href;
			}else if(href.contains("http")) {
				web=href;
			}
			
			
			}
			String[] data = new String[6];
			data[0] = name;
			data[1] = address;
			data[2] = phone;
			data[3] = email;
			data[4]= web;
			data[5]= ""+currentPage;
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
