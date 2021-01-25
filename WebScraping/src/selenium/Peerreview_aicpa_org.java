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

public class Peerreview_aicpa_org {
	static WebDriver driver;
	List<String[]> datalist = new ArrayList<>();
	String url = "https://peerreview.aicpa.org/public_file_search.html";
	public Peerreview_aicpa_org() {
		System.setProperty("webdriver.chrome.driver", "D:\\Freelancer\\thoan_excel\\chromedriver\\chromedriver.exe");
		driver = new ChromeDriver();
		driver.manage().window().maximize();
	}

	public static void main(String[] args) {

		
		String excelPath ="D:\\Freelancer\\thoan_excel\\Results\\peerreview aicpa list.xlsx";
		Peerreview_aicpa_org t = new Peerreview_aicpa_org();
//		t.processWeb(url);
		t.processExcel(excelPath);

	}
	
	public void processExcel(String filePath) {

		try {
			
			FileInputStream inputFile = new FileInputStream(new File(filePath));
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

						String state = table1Sheet.getRow(i).getCell(0).getStringCellValue();
						List<String[]> pasrelink = processWeb(state, i+1);
						
						if(pasrelink==null) {
							break;
						}else {
							datalist.addAll(pasrelink);
						}
						
						appendExcelFile("aicpa.org", "");
						System.out.println("===============Write State==================== :" + state);
						datalist.clear();
						break;
					} catch (Exception e) {
						System.out.println("EEEEEERRRRRRRRRR: " + i);
						e.printStackTrace();
					}
				}
				appendExcelFile("aicpa.org", "");
				System.out.println("===============Finish====================" );
			}
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	public List<String[]> processWeb(String state, int rowIndex) {
		List<String[]> result = new ArrayList<>();
		try {
			driver.get(url);
			Thread.sleep(1000);
			driver.switchTo().frame(driver.findElement(By.xpath("//*[@id=\"PegaGadgetIfr\"]")));
			final Select selectBox = new Select(driver.findElement(By.xpath("//*[@id=\"State\"]")));
			selectBox.selectByValue(state);
			WebElement submit = driver.findElement(By.xpath("//*[@id=\"RULE_KEY\"]/div[6]/div/div/div/div/div/div/div/div/div/div/div/div/div/div/div/span/button"));
			submit.click();
			Thread.sleep(5000);
			int pageNumber =0;
			while(true) {
				pageNumber++;
				processFrame(state,pageNumber);
				appendExcelFile("aicpa.org", "");
				datalist.clear();
				System.out.println("===========Write "+state+"===========: "+pageNumber);
				Thread.sleep(2000);
				try {
					WebElement nextPage = driver.findElement(By.xpath("//*[@id=\"grid-desktop-paginator\"]/tbody/tr/td[5]/nobr/label/a"));
					nextPage.click();
					Thread.sleep(5000);
				}catch (Exception e) {
					e.printStackTrace();
					break;
				}
			}
			
			System.out.println("===========Finish===========: "+state);
			
				
				
			
		}catch (Exception e) {
			e.printStackTrace();
			System.out.println("ERRRRRR: "+ state);
			result=null;
		}
		return result;
	}

	public void processFrame(String state, int pageNum) {
		int rowsCount = driver.findElement(By.xpath("//*[@id=\"bodyTbl_right\"]")).findElements(By.tagName("tr")).size();
		for(int i =1; i< rowsCount;i++) {
			try {
			List<WebElement> nextPages = driver.findElements(By.xpath("//*[@id=\"grid-desktop-paginator\"]/tbody/tr/td[3]/nobr/a"));
			for(WebElement nexpPage : nextPages) {
				if(nexpPage.getText().equals(""+pageNum)) {
					nexpPage.click();
					Thread.sleep(2000);
				}
			}
			List<WebElement> rows = driver.findElement(By.xpath("//*[@id=\"bodyTbl_right\"]")).findElements(By.tagName("tr"));
			List<WebElement> tds = rows.get(i).findElements(By.tagName("td"));
			WebElement link = tds.get(0).findElement(By.tagName("a"));
			String city=tds.get(2).findElement(By.tagName("span")).getText();
			String firm = link.getText();
			link.click();
			String address1 = driver.findElement(By.xpath("//*[@id=\"CT\"]/div/div/span")).getText()+", "+city+", "+ state;
			String[] data = new String[4];
			data[0] = firm;
			data[1] = address1;
			data[2] = state;
			data[3] = ""+pageNum;
			datalist.add(data);
			
			
			WebElement backBtn = driver.findElement(By.xpath("//*[@id=\"CT\"]/div/div/div/div[2]/div/div/div/div/div[1]/div/div/span/button"));
			backBtn.click();
			Thread.sleep(2000);
			
			}catch (Exception e) {
				e.printStackTrace();
				String[] data = new String[4];
				data[0] = "";
				data[1] = ""+i;
				data[2] = state;
				data[3] = ""+pageNum;
				datalist.add(data);
				System.out.println("Page Num: "+ pageNum+", Row:"+ i);
			}
		}
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
			if(!sheetName.isEmpty()) {
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
		if(!sheetName.isEmpty()) {
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
