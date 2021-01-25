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

public class Coud_withgoogle_com{
	static WebDriver driver;
	List<String[]> datalist = new ArrayList<>();
	String filename = "cloud.withgoogle.com";
	String fileResult = "cloud.withgoogle.com_result";

	public Coud_withgoogle_com() {
		System.setProperty("webdriver.chrome.driver", "D:\\Freelancer\\thoan_excel\\driver83\\chromedriver.exe");
		driver = new ChromeDriver();
		driver.manage().window().maximize();
	}

	public static void main(String[] args) {

		String url = "https://cloud.withgoogle.com/partners/";
		Coud_withgoogle_com t = new Coud_withgoogle_com();
//		t.processWeb2(url);
		 t.processExcel();

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
						String name = table1Sheet.getRow(i).getCell(0).getStringCellValue();
						String link=table1Sheet.getRow(i).getCell(1).getStringCellValue();

						processWeb(name,link,  i + 1);
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

	public void processWeb(String name, String link,int rowIndex) {
		try {
			driver.get(link);
			Thread.sleep(2000);
			WebDriverWait wait = new WebDriverWait(driver, 120);
			WebElement detailLink = wait.until(ExpectedConditions.presenceOfElementLocated(
					By.cssSelector("div[class='detail-links']")));
			List<WebElement> urlE = driver.findElements(By.cssSelector("a[class='detail-links__link']"));
			String url="";
			if(!urlE.isEmpty()) {
				url = urlE.get(0).getAttribute("href");
			}
			datalist.add(new String[] { link, name, url});

		} catch (Exception e) {
			System.out.println("ERRRRRR: " + rowIndex);
			datalist.add(new String[] { link, name, "ERROR"});
		}
	}

	public void processWeb2(String url) {
		int count=0;
		driver.get(url);
		while(true) {
		try {
			WebDriverWait wait = new WebDriverWait(driver, 120);
			WebElement loadingBtn = wait.until(ExpectedConditions.presenceOfElementLocated(
					By.xpath("//*[@id=\"load-more-cards-button\"]")));
			clickAction(loadingBtn);
			count=0;
		}catch (Exception e) {
			e.printStackTrace();
			count++;
			if(count==5) {
				break;
			}
		}
		}
		
		List<WebElement> cards = driver.findElements(By.xpath("//*[@id=\"cards__container\"]/div"));
		int i=0;
		for(WebElement card: cards) {
			String link = card.findElement(By.tagName("a")).getAttribute("href");
			datalist.add(new String[] {link});
			if(datalist.size()%50==0) {
				i++;
				appendExcelFile(filename, "");
				System.out.println("Write size="+(i*50));
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
