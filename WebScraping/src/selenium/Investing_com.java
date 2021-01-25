package selenium;

import java.awt.Color;
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
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;

public class Investing_com {
	static WebDriver driver;
	List<String[]> datalist = new ArrayList<>();
	String filename = "investing2.com";
	String fileResult = "investing.com_result2";

	public Investing_com() {
		System.setProperty("webdriver.chrome.driver", "D:\\Freelancer\\thoan_excel\\driver79\\chromedriver.exe");
		driver = new ChromeDriver();
		driver.manage().window().maximize();
	}

	public static void main(String[] args) {

		 String url = "https://www.investing.com/stock-screener/?sp=country::5%7Csector::a%7Cindustry::a%7CequityType::a%7Cexchange::1%3Ceq_market_cap;1";
		Investing_com t = new Investing_com();
//		t.processWeb2();
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

				for (int i = 400; i < table1Sheet.getPhysicalNumberOfRows(); i++) {
					try {

						Cell cell = table1Sheet.getRow(i).getCell(0);
						if (cell == null) {
							continue;
						}

						String url = table1Sheet.getRow(i).getCell(0).getStringCellValue();
						String symbol=table1Sheet.getRow(i).getCell(1).getStringCellValue();
						if(url.isEmpty()) {
							continue;
						}
						List<String[]> pasrelink = processWeb(url, symbol,i + 1);

						if (pasrelink == null) {
							break;
						} else {
							datalist.addAll(pasrelink);

						}
						
//						processWeb2(url);
						if((i+1)%50==0) {
							appendExcelFile(fileResult, "");
							System.out.println("WWWW"+i);
						}
					} catch (Exception e) {
						System.out.println("EEEEEERRRRRRRRRR: " + i+1);
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
	public void processWeb2(String isin) {
		
		
		try {
			
			WebElement searchE=driver.findElement(By.xpath("/html/body/div[5]/header/div[1]/div/div[3]/div[1]/input"));
			searchE.clear();
			searchE.sendKeys(isin);
			Thread.sleep(500);
			List<WebElement> items = driver.findElements(By.xpath("/html/body/div[5]/header/div[1]/div/div[3]/div[2]/div[1]/div[1]/div[2]/div/a"));
			if(items.isEmpty()) {
				datalist.add(new String[] {isin,""});
			}else {
				String link = items.get(0).getAttribute("href");
				datalist.add(new String[] {isin,link});
			}
		}catch (Exception e) {
			e.printStackTrace();
			System.out.println("XXXX"+isin);
		}
	}

	public List<String[]> processWeb(String url, String symbol, int rowIndex) {
		List<String[]> result = new ArrayList<>();
		try {
			driver.get(url);
			Thread.sleep(3000);
			List<WebElement> rights = driver
					.findElements(By.xpath("//*[@id=\"quotes_summary_current_data\"]/div[2]/div"));
//			String type = "";
//			String market = "";
			String isin = "";
			for (WebElement right : rights) {
				List<WebElement> spans = right.findElements(By.tagName("span"));
				String itemName = spans.get(0).getText();

//				if (itemName.equals("Type:")) {
//					type = spans.get(1).getText();
//				} else if (itemName.equals("Market:")) {
//					market = spans.get(1).getText();
//				} else 
				if (itemName.equals("ISIN:")) {
					isin = spans.get(1).getText();
				}
			}

			WebElement title = driver.findElement(By.xpath("//*[@id=\"leftColumn\"]/div[1]/h1"));
			WebElement cid = driver.findElement(By.xpath("//*[@id=\"leftColumn\"]/div[1]/div[4]"));
//			WebElement currency = driver.findElement(By.xpath("//*[@id=\"quotes_summary_current_data\"]/div[1]/div[2]/div[2]/span[4]"));

			String[] data = new String[5];
			data[0] = url;
			data[1] = title.getText();
			data[2]=  symbol;
//			data[3] = market;
			data[3] = cid.getAttribute("data-pair-id");
//			data[5] = currency.getText();
//			data[6] = type;
			data[4] = isin;
			result.add(data);

		} catch (Exception e) {
			e.printStackTrace();
			System.out.println("ERRRRRR: " + rowIndex);
			result = null;
		}
		return result;
	}

	public List<String[]> processWeb_bk(String url, int rowIndex) {
		List<String[]> result = new ArrayList<>();
		try {
			driver.get(url);
			Thread.sleep(3000);
			String market_type = driver
					.findElement(By.xpath("//*[@id=\"quotes_summary_current_data\"]/div[2]/div[1]/span[2]")).getText();
			WebElement table = driver.findElement(By.xpath("//*[@id=\"leftColumn\"]/table[3]"));
			if (table.getAttribute("class").equals("genTbl openTbl")) {
				List<WebElement> rows = driver.findElements(By.xpath("//*[@id=\"leftColumn\"]/table[3]/tbody/tr"));
				for (WebElement row : rows) {
					List<WebElement> tds = row.findElements(By.tagName("td"));
					List<WebElement> titleList = tds.get(1).findElements(By.tagName("a"));
					WebElement titleE = null;
					String currency = tds.get(7).getText();

					if (titleList.isEmpty()) {
						titleE = tds.get(1);
					} else {
						titleE = titleList.get(0);
					}
					String title = titleE.getAttribute("title");
					String[] data = new String[8];
					data[0] = title.substring(0, title.lastIndexOf('(') - 1).trim();
					data[1] = title.substring(title.lastIndexOf('(') - 1).trim();

					data[2] = titleE.getText();
					data[3] = tds.get(2).getAttribute("id");
					data[4] = currency;
					data[5] = market_type;
					data[6] = driver.getCurrentUrl();
					data[7] = "" + rowIndex;
					result.add(data);
				}
			} else {
				String[] data = new String[8];
				data[0] = "";
				data[1] = "";

				data[2] = "";
				data[3] = "";
				data[4] = "";
				data[5] = "";
				data[6] = driver.getCurrentUrl();
				data[7] = "" + rowIndex;
				result.add(data);
				System.out.println(data[6]);
			}

		} catch (Exception e) {
			e.printStackTrace();
			System.out.println("ERRRRRR: " + rowIndex);
			result = null;
		}
		return result;
	}

	public void processWeb() {
		String[] links = new String[] { 
//				"https://www.investing.com/search/?q=tundra&tab=quotes",
//				"https://www.investing.com/search/?q=aberdeen&tab=quotes",
//				"https://www.investing.com/search/?q=ansvar&tab=quotes",
//				"https://www.investing.com/search/?q=amf&tab=quotes",
//				"https://www.investing.com/search/?q=c%20worldwide&tab=quotes",
//				"https://www.investing.com/search/?q=Carnegie&tab=quotes",
				"https://www.investing.com/search/?q=Didner%20&tab=quotes",
				"https://www.investing.com/search/?q=handelsbanken&tab=quotes", 
				"https://www.investing.com/search/?q=Lannebo&tab=quotes",
//				"https://www.investing.com/search/?q=odin&tab=quotes",
				"https://www.investing.com/search/?q=priornilsson&tab=quotes",
//				"https://www.investing.com/search/?q=skagen&tab=quotes",
//				"https://www.investing.com/search/?q=skandia&tab=quotes",
//				"https://www.investing.com/search/?q=spiltan&tab=quotes",
//				"https://www.investing.com/search/?q=spp&tab=quotes" 
				};
		for (String url : links) {
			driver.get(url);
			
			try {
			Thread.sleep(2000);
			}catch (Exception e) {
				e.printStackTrace();
			}
			String resultT= driver.findElement(By.xpath("//*[@id=\"fullColumn\"]/div/div[3]/div[1]")).getText();
			int pageSize = 0;
			while (true) {
				try {
					Thread.sleep(1000);
					List<WebElement> rows = driver
							.findElements(By.xpath("//*[@id=\"fullColumn\"]/div/div[3]/div[3]/div/a")); 
					
					if (rows.size() > pageSize) {
						pageSize = rows.size();
						JavascriptExecutor jse = (JavascriptExecutor) driver;
						jse.executeScript("window.scrollBy(0,1000)");
						Thread.sleep(3000);
					} else {
						System.out.println("===Size ="+rows.size()+", "+resultT);
						break;
					}
					

				} catch (Exception e) {
					e.printStackTrace();
					break;
				}
			}
			
			processFrame();
			appendExcelFile(filename, "");
		}
		System.out.println("===============Finish====================");
	}
	
	public void processWeb2() {
		String[] links = new String[] { 
				"https://www.investing.com/stock-screener/?sp=country::5|sector::a|industry::a|equityType::a|exchange::2%3Ceq_market_cap;1"
				};
		for (String url : links) {
			driver.get(url);

			while (true) {
				try {
					Thread.sleep(1000);
					if(processFrame2()) {
						appendExcelFile(filename, "");
					}else {
						break;
					}
					List<WebElement> nextBtns=driver.findElements(By.xpath("//*[@id=\"paginationWrap\"]/div[3]/a"));
					if(nextBtns.isEmpty()) {
						break;
					}else {
						String nextItems=nextBtns.get(0).getAttribute("title");
						System.out.println(nextItems);
						clickAction(nextBtns.get(0));
						Thread.sleep(2000);
					}
							
				} catch (Exception e) {
					e.printStackTrace();
					break;
				}
			}
			
		}
		System.out.println("===============Finish====================");
	}
	
	public boolean processFrame2() {
		try {
		List<WebElement> rows = driver
				.findElements(By.xpath("//*[@id=\"resultsTable\"]/tbody/tr")); 
		
		for (WebElement row : rows) {
			List<WebElement> tds= row.findElements(By.tagName("td"));
			String link = tds.get(1).findElement(By.tagName("a")).getAttribute("href");
			String symbol = tds.get(2).getText();
			datalist.add(new String[] {link, symbol});
		}
		}catch (Exception e) {
			e.printStackTrace();
			return false;
		}
		return true;
	}


	public void processFrame() {

		List<WebElement> rows = driver.findElements(By.xpath("//*[@id=\"fullColumn\"]/div/div[3]/div[3]/div/a"));
		for (WebElement row : rows) {
			String link = row.getAttribute("href");
			String[] data = new String[1];
			data[0] = link;

			datalist.add(data);
		}
	}
	
	public void clickAction(WebElement element) {
		JavascriptExecutor js = (JavascriptExecutor)driver;
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
			File file = new File("D:\\Freelancer\\thoan_excel\\Results\\" + fileName + ".xlsx");
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
				XSSFCellStyle style = (XSSFCellStyle) workbook.createCellStyle();
		        XSSFColor myColor = new XSSFColor(Color.BLUE);
		        style.setFillForegroundColor(myColor);
		        style.setFillBackgroundColor(myColor);
				sheet.getRow(rowNum-1).getCell(0).setCellStyle(style);
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
