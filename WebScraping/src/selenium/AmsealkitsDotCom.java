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

public class AmsealkitsDotCom {
	static WebDriver driver;
	List<String[]> datalist = new ArrayList<>();
	String fileName = "amseal.com.xlsx";
	String fileResult = "amseal.com_result.xlsx";

	public AmsealkitsDotCom() {
		System.setProperty("webdriver.chrome.driver", "D:\\Freelancer\\thoan_excel\\chromedriver\\chromedriver.exe");
		driver = new ChromeDriver();
		driver.manage().window().maximize();
	}

	public static void main(String[] args) {

		String url = "http://amseal.com/index_files/Page1368.htm";
		AmsealkitsDotCom t = new AmsealkitsDotCom();
//		t.processWeb(url);
		t.processExcel2();

	}


	public void processWeb(String url) {

		try {
			driver.get(url);
			Thread.sleep(3000);
			String c1 = driver.findElement(By.xpath("/html/body/div[1]/span[21]/table/tbody/tr/td/div/h2")).getText();
			
			List<WebElement> as = driver.findElement(By.xpath("/html/body/div[1]/span[21]/table/tbody/tr/td/div")).findElements(By.tagName("a"));
			for(WebElement a : as) {
				
					
						String link = a.getAttribute("href");
						String c2=a.getText();
						datalist.add(new String[] {link, c1, c2});
						
					
				
			}
			
		} catch (Exception e) {
			e.printStackTrace();
		}
		appendExcelFile(fileName, "");
		System.out.println("===============Finish====================");
	}

	

	public void processExcel() {

		try {
			FileInputStream inputFile = new FileInputStream(
					new File("D:\\Freelancer\\thoan_excel\\Results\\" + fileName));
			XSSFWorkbook wb = new XSSFWorkbook(inputFile);
			XSSFSheet table1Sheet = wb.getSheetAt(1);

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
						String c1 = table1Sheet.getRow(i).getCell(1).getStringCellValue();
						
						if(!processWeb(link, c1, (i+1))) {
							
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

	public boolean processWeb(String c2Link,String c1, int rowIndex) {

		try {
			driver.get(c2Link);
			Thread.sleep(3000);
			String c2 = driver.findElement(By.xpath("/html/body/div[1]/span[7]/table/tbody/tr/td/div")).getText();
			List<WebElement> items =driver.findElements(By.xpath("/html/body/div[1]/table[1]/tbody/tr")); 
			for(int i=0; i<items.size();i++) {
				List<WebElement> tds = items.get(i).findElements(By.tagName("td"));
				String att1=tds.get(2).getText();
				String att2 ="";
				if(tds.size()>3) {
					att2 = tds.get(3).getText();
				}
				List<WebElement> c3Es = tds.get(0).findElements(By.tagName("a"));
				String c3="";
				String c3Link="";
				if(!c3Es.isEmpty()) {
					for(WebElement c3E: c3Es) {
						c3 = c3E.getText();
						c3Link = c3E.getAttribute("href");
						String[] data = new String[7];
						data[0]=c1;
						data[1]=c2;
						data[2]=c3;
						data[3]=att1;
						data[4]=att2;
						data[5]=c3Link;
						data[6]=c2Link;
						datalist.add(data);
					}
					
				
				}else {
					c3= tds.get(0).getText();
					String[] data = new String[7];
					data[0]=c1;
					data[1]=c2;
					data[2]=c3;
					data[3]=att1;
					data[4]=att2;
					data[5]=c3Link;
					data[6]=c2Link;
					datalist.add(data);
				}
				
			}
			
		} catch (Exception e) {
			e.printStackTrace();
			System.out.println("ERRRRRR: " + rowIndex);
			return false;
			
		}
		return true;

	}
	
	public void processExcel2() {

		try {
			FileInputStream inputFile = new FileInputStream(
					new File("D:\\Freelancer\\thoan_excel\\Results\\" + fileName));
			XSSFWorkbook wb = new XSSFWorkbook(inputFile);
			XSSFSheet table1Sheet = wb.getSheetAt(3);

			if (table1Sheet == null) {
				System.out.println("KO CO SHEET :" + table1Sheet);
			} else {

				for (int i = 334; i < table1Sheet.getPhysicalNumberOfRows(); i++) {
					try {

						Cell cell = table1Sheet.getRow(i).getCell(0);
						if (cell == null) {
							continue;
						}

						String c3Link = table1Sheet.getRow(i).getCell(5).getStringCellValue();
						String c1 = table1Sheet.getRow(i).getCell(0).getStringCellValue();
						String c2 = table1Sheet.getRow(i).getCell(1).getStringCellValue();
						String c3 = table1Sheet.getRow(i).getCell(2).getStringCellValue();
						String att1 = table1Sheet.getRow(i).getCell(3).getStringCellValue();
						String att2 = table1Sheet.getRow(i).getCell(4).getStringCellValue();
						String c2Link = table1Sheet.getRow(i).getCell(6).getStringCellValue();
						if(c3Link.isEmpty()) {
							continue;
						}
						if(!processWeb(c3Link, c1,c2, c3, att1, att2,c2Link, (i+1))) {
							
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
	
	public boolean processWeb(String c3Link, String c1, String c2, String c3, String att1, String att2, String c2Link, int rowIndex) {

		try {
			
			driver.get(c3Link);
			Thread.sleep(3000);
			String c3_2 = driver.findElement(By.xpath("/html/body/center/h1")).getText();
			List<WebElement> items =driver.findElements(By.xpath("/html/body/center/table/tbody/tr")); 
			for(int i=1; i<items.size();i++) {
				List<WebElement> tds = items.get(i).findElements(By.tagName("td"));
				String name=tds.get(0).getText();
				String description = tds.get(1).getText();
				String listPrice = tds.get(2).getText();
				String yourPrice = tds.get(3).getText();
					String[] data = new String[11];
					data[0]=c1;
					data[1]=c2;
					data[2]=c3;
					data[3]=c3_2;
					data[3]=att1;
					data[4]=att2;
					data[5]=name;
					data[6]=description;
					data[7]=listPrice;
					data[8]=yourPrice;
					data[9]=c3Link;
					data[10]=c2Link;
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
