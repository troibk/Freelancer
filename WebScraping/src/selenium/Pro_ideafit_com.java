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

public class Pro_ideafit_com {
	static WebDriver driver;
	List<String[]> datalist = new ArrayList<>();
	
	String excelPath ="D:\\Freelancer\\thoan_excel\\Results\\";
	String filename="pro.ideafit.com.xlsx";
	public Pro_ideafit_com() {
		System.setProperty("webdriver.chrome.driver", "D:\\Freelancer\\thoan_excel\\chromedriver\\chromedriver.exe");
		driver = new ChromeDriver();
		driver.manage().window().maximize();
	}

	public static void main(String[] args) {
		String url = "https://pro.ideafit.com/find-personal-trainer";
		Pro_ideafit_com t = new Pro_ideafit_com();
		t.processWeb(url);
//		t.processExcel();

	}
	
	public void processExcel() {

		try {
			String filePath= excelPath+filename;
			FileInputStream inputFile = new FileInputStream(new File(filePath));
			XSSFWorkbook wb = new XSSFWorkbook(inputFile);
			XSSFSheet table1Sheet = wb.getSheetAt(1);

			if (table1Sheet == null) {
				System.out.println("KO CO SHEET :" + table1Sheet);
			} else {

				for (int i = 15; i < table1Sheet.getPhysicalNumberOfRows(); i++) {
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
						
						appendExcelFile("");
						System.out.println("===============Write State==================== :" + state);
						datalist.clear();
				
					} catch (Exception e) {
						System.out.println("EEEEEERRRRRRRRRR: " + i);
						e.printStackTrace();
					}
				}
				appendExcelFile("");
				System.out.println("===============Finish====================" );
			}
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	public void processWeb(String url) {
		List<String> states = new ArrayList<>();
		try {
			driver.get(url);
			Thread.sleep(5000);
			List<WebElement> rowsList= driver.findElements(By.cssSelector("div[class='row col-links']"));
			for(int i =1; i<rowsList.size();i++) {
				List<WebElement> links = rowsList.get(i).findElements(By.tagName("a"));
				for(WebElement linkE : links) {
					String link = linkE.getAttribute("href");
					states.add(link);
				}
			}
			
			for(String state : states) {
				driver.get(state);
				Thread.sleep(3000);
				List<WebElement> rowsList2= driver.findElements(By.cssSelector("div[class='row col-links']"));
				for (WebElement row : rowsList2) {
					List<WebElement> links = row.findElements(By.tagName("a"));
					for(WebElement linkE : links) {
						String link = linkE.getAttribute("href");
						datalist.add(new String[] {link});
					}
				}
			}
			
			appendExcelFile("");
			
		}catch (Exception e) {
			e.printStackTrace();
		}

	}
	
	public List<String[]> processWeb(String url, int rowIndex) {
		List<String[]> result = new ArrayList<>();
		try {
			driver.get(url);
			Thread.sleep(1000);
				
			
		}catch (Exception e) {
			e.printStackTrace();
			result=null;
		}
		return result;
	}

	public void processFrame(String state, int pageNum) {
		List<WebElement> rows = driver.findElement(By.xpath("//*[@id=\"bodyTbl_right\"]")).findElements(By.tagName("tr"));
		for(int i =1; i< rows.size();i++) {
			try {
			List<WebElement> tds = rows.get(i).findElements(By.tagName("td"));
			String firm = tds.get(0).findElement(By.tagName("a")).getText();
			String city=tds.get(2).findElement(By.tagName("span")).getText();
			String[] data = new String[4];
			data[0] = firm;
			data[1] = city;
			data[2] = state;
			data[3] = ""+pageNum;
			datalist.add(data);
			}catch (Exception e) {
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

	private void appendExcelFile(String sheetName) {
		Workbook workbook = null;
		Sheet sheet;
		try {
			File file = new File("D:\\Freelancer\\thoan_excel\\Results\\"+filename);
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
