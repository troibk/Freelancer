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

public class Equityresearch_com {
	static WebDriver driver;
	List<String[]> datalist = new ArrayList<>();
	String filename = "equity-research.com";
	String fileResult = "equity-research.com_result";

	public Equityresearch_com() {
		System.setProperty("webdriver.chrome.driver", "D:\\Freelancer\\thoan_excel\\driver81_2\\chromedriver.exe");
		driver = new ChromeDriver();
		driver.manage().window().maximize();
	}

	public static void main(String[] args) {

		String url = "http://equity-research.com/list-of-top-200-investment-banks-and-boutiques/";
		Equityresearch_com t = new Equityresearch_com();
		t.processWeb2(url);
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

	public void processWeb2(String url) {
		try {
			driver.get(url);
			Thread.sleep(1000);
			List<WebElement> items = driver.findElements(By.xpath("//*[@id=\"post-822\"]/div/p"));
			System.out.println(items.size());
			for(WebElement item : items) {
				String name = item.getText();
				List<WebElement> as = item.findElements(By.tagName("a"));
				String web="";
				String mail="";
				if(as.size()>0) {
					web= as.get(0).getAttribute("href");
				}
				if(as.size()==2) {
					mail = as.get(1).getAttribute("href");
				}
				datalist.add(new String[] {web, mail, name});
				if(datalist.size()==50) {
					appendExcelFile(filename, "");
				}
			}
		}catch (Exception e) {
			e.printStackTrace();
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
