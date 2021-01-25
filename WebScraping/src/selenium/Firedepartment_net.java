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
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;

public class Firedepartment_net {
	static WebDriver driver;
	List<String[]> datalist = new ArrayList<>();
	String fileName = "firedepartment.net.xlsx";
	String fileResult = "firedepartment.net_result.xlsx";

	public Firedepartment_net() {
		System.setProperty("webdriver.chrome.driver", "D:\\Freelancer\\thoan_excel\\chromedriver\\chromedriver.exe");
		driver = new ChromeDriver();
		driver.manage().window().maximize();
	}

	public static void main(String[] args) {

		String url = "https://www.firedepartment.net/";
		Firedepartment_net t = new Firedepartment_net();
//		t.processWeb(url);
		 t.processExcel();

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

	public void processWeb(String url) {

		try {
			driver.get(url);
			Thread.sleep(5000);
			List<WebElement> rows = driver.findElements(By.xpath("/html/body/div/div[2]/div[2]/div[2]/div[1]/div/div[2]/div/div[1]/ul/li"));
			for(WebElement row: rows) {
				String link = row.findElement(By.tagName("a")).getAttribute("href");
				String state= row.getText();
				datalist.add(new String[] {link, state});
				
			}
			rows = driver.findElements(By.xpath("/html/body/div/div[2]/div[2]/div[2]/div[1]/div/div[2]/div/div[2]/ul/li"));
			for(WebElement row: rows) {
				String link = row.findElement(By.tagName("a")).getAttribute("href");
				String state= row.getText();
				datalist.add(new String[] {link, state});
			}
			rows = driver.findElements(By.xpath("/html/body/div/div[2]/div[2]/div[2]/div[1]/div/div[2]/div/div[3]/ul/li"));
			for(WebElement row: rows) {
				String link = row.findElement(By.tagName("a")).getAttribute("href");
				String state= row.getText();
				datalist.add(new String[] {link, state});
			}
			appendExcelFile(fileName, "");
		} catch (Exception e) {
			e.printStackTrace();
		}
		System.out.println("===============Finish====================");
	}
	int currentIndex=0;
	public int processPage() {
		List<WebElement> rows = driver.findElements(By.xpath("//*[@id=\"search-results\"]/a"));
		for (int i =0;i<rows.size();i++) {
			currentIndex++;
			String link = rows.get(i).getAttribute("href");
			String name = rows.get(i).findElement(By.xpath("//*[@id=\"search-results\"]/a["+(i+1)+"]/div/h4/div[1]")).getText();
			String details = rows.get(i).findElement(By.cssSelector("div[class='details']")).getText();
			String[] data = new String[4];
			data[0] = ""+currentIndex;
			data[1] = name;
			data[2] = details;
			data[3] = link;
			datalist.add(data);
		}
		return rows.size();
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
						String name = table1Sheet.getRow(i).getCell(1).getStringCellValue();
						String state = table1Sheet.getRow(i).getCell(2).getStringCellValue();
						if(processWeb(link, name, state, (i+1))){
							appendExcelFile(fileResult, "");
						}else {
							break;
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

	public boolean processWeb(String link,String name, String state, int rowIndex) {
		
		try {
			driver.get(link);
			Thread.sleep(3000);
			List<WebElement> rows = driver.findElement(By.xpath("/html/body/section[3]/div/div/div[1]/div[2]/div/ul")).findElements(By.tagName("a"));
			for(int i=0;i<rows.size();i++) {
				String url = rows.get(i).getAttribute("href");
				String name2=rows.get(i).getText();
				datalist.add(new String[] {url, name2, name,state});
			}

		} catch (Exception e) {
			e.printStackTrace();
			System.out.println("ERRRRRR: " + rowIndex);
			return false;
		}
		return true;
	}
	
public boolean processWeb2(String link,String name2, String name, String state, int rowIndex) {
		
		try {
			driver.get(link);
			Thread.sleep(3000);
			List<WebElement> rows = driver.findElements(By.xpath("//*[@id=\"content\"]/div[1]/div[1]/div/div[3]/div[1]/div/div[2]/address[1]"));
			String address=rows.isEmpty()? "": rows.get(0).getText();
			datalist.add(new String[] {address, name2, name, state});
		} catch (Exception e) {
			e.printStackTrace();
			System.out.println("ERRRRRR: " + rowIndex);
			return false;
		}
		return true;
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
