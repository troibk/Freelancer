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

public class lsba_org {
	static WebDriver driver;
	List<String[]> datalist = new ArrayList<>();
	String filename = "isba.org";
	String fileResult = "isba.org_result";

	public lsba_org() {
		System.setProperty("webdriver.chrome.driver", "D:\\Freelancer\\thoan_excel\\driver79\\chromedriver.exe");
		driver = new ChromeDriver();
		driver.manage().window().maximize();
	}

	public static void main(String[] args) {

		String url = "https://www.lsba.org/Public/MembershipDirectory.aspx";
		lsba_org t = new lsba_org();
		t.processWeb2();
//		 t.processExcel();

	}

	public void processExcel() {

		try {
			String filePath = "D:\\Freelancer\\thoan_excel\\Results\\" + filename + ".xlsx";
			FileInputStream inputFile = new FileInputStream(new File(filePath));
			XSSFWorkbook wb = new XSSFWorkbook(inputFile);
			XSSFSheet table1Sheet = wb.getSheetAt(0);

			if (table1Sheet == null) {
				System.out.println("KO CO SHEET :" + table1Sheet);
			} else {
				appendExcelFile(fileResult, "");
				for (int i = 0; i < table1Sheet.getPhysicalNumberOfRows(); i++) {
					try {

						Cell cell = table1Sheet.getRow(i).getCell(0);
						if (cell == null) {
							continue;
						}
						String email=table1Sheet.getRow(i).getCell(4).getStringCellValue();
						if(!email.isEmpty()) {
							continue;
						}
						String name = table1Sheet.getRow(i).getCell(1).getStringCellValue();
						String link=table1Sheet.getRow(i).getCell(2).getStringCellValue();
						String state = table1Sheet.getRow(i).getCell(3).getStringCellValue();
						

						processWeb(name,link, state,  i + 1);
						if ((i + 1) % 50 == 0) {
							appendExcelFile(fileResult, "");
							System.out.println("WWWW" + i);
						}
					} catch (Exception e) {
						System.out.println("EEEEEERRRRRRRRRR: " + i + 1);
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

	public void processWeb(String name, String link, String state, int rowIndex) {
		try {
			driver.get(link);
			Thread.sleep(5000);
			List<WebElement> emailE = driver.findElements(By.xpath("//*[@id=\"Main_Content_Placeholder_C004_lnkEmailAttorney\"]"));
			String email="";
			if(!emailE.isEmpty()) {
				email = emailE.get(0).getAttribute("href");
			}
			
			List<WebElement> firmE = driver.findElements(By.xpath("//*[@id=\"Main_Content_Placeholder_C004_pnlCompany\"]"));
			String firm="";
			if(!firmE.isEmpty()) {
				firm = firmE.get(0).getText();
			}
			datalist.add(new String[] { link, firm, name, email, state });

		} catch (Exception e) {
			System.out.println("ERRRRRR: " + rowIndex);
			datalist.add(new String[] { link, "", name, "", state });
		}
	}

	public void processWeb2() {
		String ori_url = "https://www.lsba.org/Public/MembershipDirectory.aspx";
		String[] keys = new String[] {"A", "B","C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };
		for (String key : keys) {
			driver.get(ori_url);
			WebDriverWait wait = new WebDriverWait(driver, 60);
			WebElement lastNameE = wait.until(ExpectedConditions.presenceOfElementLocated(
					By.xpath("//*[@id=\"TextBoxLastName\"]")));
			lastNameE.sendKeys(key);
			WebElement searchE = driver.findElement(By.xpath("//*[@id=\"ButtonSearch\"]"));
			clickAction(searchE);
			int pageId=0;
			while (true) {
				try {
					pageId++;
					List<String[]> datalist2 = new ArrayList<>();
					WebDriverWait wait2 = new WebDriverWait(driver, 60);
					WebElement table = wait2.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//*[@id=\"ListView1_itemPlaceholderContainer\"]")));
					List<WebElement> items = table.findElements(By.tagName("div"));
					int i=0;
					for(WebElement item : items) {
						i++;
						String eligible = table.findElement(By.xpath(".//div["+i+"]/span")).getText();
						if(eligible.contains("Not")) {
							continue;
						}
						String name = item.findElement(By.tagName("b")).getText();
						String link= item.findElement(By.tagName("a")).getAttribute("href");
						datalist2.add(new String[] { key, name, link});
					}
					JavascriptExecutor js = (JavascriptExecutor) driver;
					for(String[] data : datalist2) {
						
						js.executeScript(data[2]);
						WebDriverWait wait3 = new WebDriverWait(driver, 60);
						WebElement detailsE = wait3.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//*[@id=\"divdetails\"]")));
						
						String details = detailsE.getText();
						if(details.isEmpty()) {
							Thread.sleep(1000);
							detailsE = driver.findElement(By.xpath("//*[@id=\"divdetails\"]"));
							details= detailsE.getText();
						}
						if(details.isEmpty()) {
							Thread.sleep(2000);
							detailsE = driver.findElement(By.xpath("//*[@id=\"divdetails\"]"));
							details= detailsE.getText();
						}
						
						datalist.add(new String[] {data[0], data[1],details});
						backAction();
						Thread.sleep(500);
					}
					
					appendExcelFile(filename, "");

					System.out.println("=====finished:" + key + "_" + pageId);

					WebElement nextBtn = driver.findElement(By.xpath("//*[@id=\"DataPager1\"]/input[2]"));
					if(nextBtn.getAttribute("disabled")==null) {
						clickAction(nextBtn);
					}else {
						break;
					}
				} catch (Exception e) {
					e.printStackTrace();
					break;
				}
			}

		}
		System.out.println("===============Finish====================");
	}

	public void clickAction(WebElement element) {
		JavascriptExecutor js = (JavascriptExecutor) driver;
		js.executeScript("arguments[0].click();", element);
	}
	
	public void backAction() {
		JavascriptExecutor js = (JavascriptExecutor) driver;
		js.executeScript("window.history.go(-1)");
	}

	public void selectComboValue(final String elementName, final String value) {
		final Select selectBox = new Select(driver.findElement(By.cssSelector(elementName)));
		selectBox.selectByValue(value);
	}

	private void appendExcelFile(String fileName, String sheetName) {
		Workbook workbook = null;
		Sheet sheet;
		try {
			File file = new File("D:\\Freelancer\\thoan_excel\\Results\\" + fileName + ".xlsx");
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

}
