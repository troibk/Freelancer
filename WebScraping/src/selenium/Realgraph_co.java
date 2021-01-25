package selenium;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

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
import org.openqa.selenium.support.ui.Select;

public class Realgraph_co {
	static WebDriver driver;
	List<String[]> datalist = new ArrayList<>();
	List<Map<String, List<String>>> datalist2 = new ArrayList<>();
	String fileName = "realgraph.co.xlsx";
	String fileResult = "realgraph.co_result_2.xlsx";

	public Realgraph_co() {
		System.setProperty("webdriver.chrome.driver", "D:\\Freelancer\\thoan_excel\\driver\\chromedriver.exe");
		driver = new ChromeDriver();
		driver.manage().window().maximize();
	}

	public static void main(String[] args) {

		String url = "https://realgraph.co/transactions/financing";
		Realgraph_co t = new Realgraph_co();
//		t.processWeb(url);
		 t.processExcel();

	}

	public void processWeb(String url) {

		try {
			driver.get(url);
			Thread.sleep(3000);

			while (true) {

				String pageName = processPage();

				List<WebElement> pagingBtn = driver.findElements(By.cssSelector("a[class='pagination-right']"));
				if (pagingBtn.isEmpty()) {
					break;
				}else {
					clickAction(pagingBtn.get(0));
					Thread.sleep(3000);
				}
				
				appendExcelFile(fileName, "");
				System.out.println("============Write : " + pageName);
			}

		} catch (Exception e) {
			e.printStackTrace();
		}
		System.out.println("===============Finish====================");
	}

	public String processPage() {
		WebElement pageE = driver.findElement(By.cssSelector("p[class='search-result-pagination']"));
		String page = pageE.getText();
		List<WebElement> rows = driver.findElements(By.xpath("//*[@id=\"content\"]/div[2]/div[2]/div"));
		for (WebElement r : rows) {
			String link = r.findElement(By.tagName("a")).getAttribute("href");

			String[] data = new String[2];
			data[0] = link;
			data[1] = page;
			datalist.add(data);
		}
		return page;
	}
	
	public void clickAction(WebElement element) {
		JavascriptExecutor js = (JavascriptExecutor) driver;
		js.executeScript("arguments[0].click();", element);
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

				for (int i = 0; i < table1Sheet.getPhysicalNumberOfRows(); i++) {
					try {

						Cell cell = table1Sheet.getRow(i).getCell(0);
						if (cell == null) {
							continue;
						}

						String link = table1Sheet.getRow(i).getCell(0).getStringCellValue();
						if(table1Sheet.getRow(i).getCell(1)==null) {
							datalist.add(new String[] {link, "","","","","","",""});
							continue;
						}
						String address = table1Sheet.getRow(i).getCell(1).getStringCellValue();
						
						processWeb2(link, address,i+1);
						

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
	
	public void processWeb2(String link, String address, int rowIndex) {
		try {
			driver.get(link);
			Thread.sleep(1000);
			List<WebElement> items = driver.findElements(By.cssSelector("span[class='card-data']"));
			for (WebElement item : items) {
				List<WebElement> nameEs = item.findElements(By.tagName("span"));
				if(nameEs.isEmpty()) {
					continue;
				}
				String name = nameEs.get(0).getText().toUpperCase();
				if (name.contains(address.toUpperCase())) {
					clickAction(item);
					Thread.sleep(1000);
					List<WebElement> item2s = driver.findElement(By.xpath("//*[@id=\"content\"]/div[2]/div[1]/div[2]/div[2]")).findElements(By.cssSelector("span[class='card-data']"));
					String[] data = new String[8];
					data[0]=link;
					data[1]=address;
					
					for (int i=0;i<item2s.size();i++) {
						String value = item2s.get(i).getText();
						data[i+2]=value;
					}
					datalist.add(data);
					break;
				} 
			}
			
			if (rowIndex % 50 == 0) {
				appendExcelFile(fileResult, "");
				System.out.println("===============Write row:" + (rowIndex));
     		}

		} catch (Exception e) {
			e.printStackTrace();
			System.out.println("ERRRRRR: " + rowIndex);
			
		}
		
	}

	public Map<String, List<String>> processWeb(String link, int rowIndex) {

		Map<String, List<String>> data = new HashMap<>();
		data.put("LINK", new ArrayList<>());
		data.get("LINK").add(link);
		data.put("DATE", new ArrayList<>());
//		data.put("STRUCTURE", new ArrayList<>());
//		data.put("AMOUNT", new ArrayList<>());
//		data.put("TERM YEARS", new ArrayList<>());
//		data.put("INTEREST RATE", new ArrayList<>());
//		data.put("FIXED VS FLOATING", new ArrayList<>());
//		data.put("FINANCING TYPES", new ArrayList<>());
//		data.put("ADDRESS", new ArrayList<>());
//		data.put("BORROWER", new ArrayList<>());
//		data.put("BORROWER_REPRESENTATIVE_O", new ArrayList<>());
//		data.put("LENDER", new ArrayList<>());
//		data.put("BORROWER_REPRESENTATIVE_P", new ArrayList<>());
		String value = "";
		try {
			driver.get(link);
			Thread.sleep(500);
			WebElement dateE = driver.findElement(By.xpath("//*[@id=\"content\"]/div[2]/div[1]/div[1]/span[3]"));
			value=dateE.getText();
			data.get("DATE").add(value);
//			List<WebElement> items = driver.findElements(By.cssSelector("span[class='card-data']"));
//			for (WebElement item : items) {
//				List<WebElement> nameEs = item.findElements(By.tagName("span"));
//				if(nameEs.isEmpty()) {
//					continue;
//				}
//				String name = nameEs.get(0).getText().toUpperCase();
//				if(name.equals("BORROWER'S REPRESENTATIVE")) {
//					if(nameEs.get(0).findElements(By.tagName("div")).isEmpty()) {
//						name="BORROWER_REPRESENTATIVE_P";
//					}else {
//						name="BORROWER_REPRESENTATIVE_O";
//					}
//				}
//				if (name.contains(",")) {
//					value = name;
//					data.get("ADDRESS").add(value);
//				} else {
//					WebElement valueE = item.findElement(By.tagName("h3"));
//					value = valueE.getText();
//					if(data.get(name) !=null) {
//						data.get(name).add(value);
//					}
//				}
//			}
			
			datalist2.add(data);
			if (rowIndex % 50 == 0) {
				appendExcelFile2(fileResult, "");
				System.out.println("===============Write row:" + (rowIndex + 1));
     		}

		} catch (Exception e) {
			e.printStackTrace();
			System.out.println("ERRRRRR: " + rowIndex);
			return null;
		}
		return data;
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
			int cellNum=0;
			if (!sheetName.isEmpty()) {
				sheet.createRow(rowNum++).createCell(cellNum).setCellValue(sheetName);
			}
			
			for (Map<String, List<String>> d : datalist2) {
				Row row = sheet.createRow(rowNum++);
				
				List<String> LINK = d.get("LINK");
				String link =LINK.get(0);
				writeExcelRow(row, cellNum, LINK, link);
				
				cellNum++;
				List<String> DATE = d.get("DATE");
				writeExcelRow(row, cellNum, DATE, link);
				
//				cellNum++;
//				List<String> STRUCTURE = d.get("STRUCTURE");
//				writeExcelRow(row, cellNum, STRUCTURE, link);
//				
//				cellNum++;
//				List<String> AMOUNT = d.get("AMOUNT");
//				writeExcelRow(row, cellNum, AMOUNT, link);
//				
//				cellNum++;
//				List<String> TERM_YEARS = d.get("TERM YEARS");
//				writeExcelRow(row, cellNum, TERM_YEARS, link);
//				
//				cellNum++;
//				List<String> INTEREST_RATE = d.get("INTEREST RATE");
//				writeExcelRow(row, cellNum, INTEREST_RATE, link);
//				
//				cellNum++;
//				List<String> FIXED_VS_FLOATING = d.get("FIXED VS FLOATING");
//				writeExcelRow(row, cellNum, FIXED_VS_FLOATING, link);
//				
//				cellNum++;
//				List<String> FINANCING_TYPES = d.get("FINANCING TYPES");
//				writeExcelRow(row, cellNum, FINANCING_TYPES, link);
//				
//				cellNum++;
//				List<String> ADDRESS = d.get("ADDRESS");
//				writeExcelRow(row, cellNum, ADDRESS, link);
//				
//				cellNum=cellNum+13;
//				List<String> BORROWER = d.get("BORROWER");
//				writeExcelRow(row, cellNum, BORROWER, link);
//				
//				cellNum=cellNum+10;
//				List<String> BORROWER_REPRESENTATIVE_O = d.get("BORROWER_REPRESENTATIVE_O");
//				writeExcelRow(row, cellNum, BORROWER_REPRESENTATIVE_O, link);
//				
//				cellNum=cellNum+10;
//				List<String> LENDER = d.get("LENDER");
//				writeExcelRow(row, cellNum, LENDER, link);
//				
//				cellNum=cellNum+10;
//				List<String> BORROWER_REPRESENTATIVE_P = d.get("BORROWER_REPRESENTATIVE_P");
//				writeExcelRow(row, cellNum, BORROWER_REPRESENTATIVE_P, link);
				cellNum=0;
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

	private void writeExcelRow(Row row, int cellNum, List<String> data, String link) {
		for (String d : data) {
			row.createCell(cellNum++).setCellValue(d);
		}
		if(data.size()>13) {
			System.out.println("XXXXXXXXXXXX:"+ link);
		}
	}
}
