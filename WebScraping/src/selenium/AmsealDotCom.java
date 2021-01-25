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

public class AmsealDotCom {
	static WebDriver driver;
	List<String[]> datalist = new ArrayList<>();
	String fileName = "amsealkits.com.xlsx";
	String fileResult = "amsealkits.com_result.xlsx";

	public AmsealDotCom() {
		System.setProperty("webdriver.chrome.driver", "D:\\Freelancer\\thoan_excel\\chromedriver\\chromedriver.exe");
		driver = new ChromeDriver();
		driver.manage().window().maximize();
	}

	public static void main(String[] args) {

		String url = "http://amsealkits.com/index_files/manufacturerbrands.htm";
		AmsealDotCom t = new AmsealDotCom();
//		t.processWeb(url);
		t.processExcel();

	}


	public void processWeb(String url) {

		try {
			driver.get(url);
			Thread.sleep(3000);
			
			List<WebElement> ps = driver.findElements(By.xpath("/html/body/div[1]/span[70]/table/tbody/tr/td/div/p"));
			for(WebElement p : ps) {
				List<WebElement> as = p.findElements(By.tagName("a"));
				if(!as.isEmpty()) {
					for(WebElement a : as) {
						String link = a.getAttribute("href");
						String brand=a.getText();
						datalist.add(new String[] {link, brand});
						
					}
				}
			}
			
		} catch (Exception e) {
			e.printStackTrace();
		}
		appendExcelFile(fileName, "");
		System.out.println("===============Finish====================");
	}

	

	public void processExcel() {
;
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

						String link = table1Sheet.getRow(i).getCell(0).getStringCellValue();
						String brand = table1Sheet.getRow(i).getCell(1).getStringCellValue();
						
						if(!processWeb(link, brand, (i+1))) {
							
							break;
						}else {
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

	public boolean processWeb(String link,String brand, int rowIndex) {

		try {
			driver.get(link);
			Thread.sleep(3000);
			List<WebElement> items =driver.findElements(By.xpath("/html/body/center[3]/table/tbody/tr")); 
			for(int i=1; i<items.size();i++) {
				List<WebElement> tds = items.get(i).findElements(By.tagName("td"));
				String name =tds.get(0).getText();
				String description=tds.get(1).getText();
				String price = tds.get(2).getText();
				String[] data = new String[5];
				data[0]=brand;
				data[1]=name;
				data[2]=description;
				data[3]=price;
				data[4]=link;
				datalist.add(data);
			}
			
		} catch (Exception e) {
			e.printStackTrace();
			System.out.println("ERRRRRR: " + rowIndex);
			return false;
			
		}
		return true;

	}
	
	public void clickAction(WebElement element) {
		JavascriptExecutor js = (JavascriptExecutor)driver;
		js.executeScript("arguments[0].click();", element);
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
