package selenium;

import java.awt.event.FocusAdapter;
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
import org.openqa.selenium.support.ui.Select;

public class Morningstar_se {
	static WebDriver driver;
	List<String[]> datalist = new ArrayList<>();
	String fileName = "investing.com.xlsx";
	String fileResult="investing.com_result.xlsx";

	public Morningstar_se() {
		System.setProperty("webdriver.chrome.driver", "D:\\Freelancer\\thoan_excel\\driver79\\chromedriver.exe");
		driver = new ChromeDriver();
		driver.manage().window().maximize();
	}

	public static void main(String[] args) {

		String url = "https://www.morningstar.se/guide/quickrank?sort=name&sortdir=asc&maincategory=&mscategory=&company=&freetext=&categorize=false&ppm=true&ppm=false&fivestar=false&tr=TrailingReturns.ThreeMonth&stdev=StandardDeviations.ThreeYear";
		Morningstar_se t = new Morningstar_se();
//		t.processWeb(url);
		t.processExcel();

	}

	public void processWeb(String url) {
		try {
			driver.get(url);
			
			while(true) {
				Thread.sleep(3000);
				String currentPage=driver.findElement(By.xpath("//*[@id=\"msse\"]/div/form/ul")).findElement(By.cssSelector("li[class='active']")).getText();
				if(processPage()) {
					appendExcelFile(fileName, "");
					System.out.println("===============Write page:" +currentPage);
				}
				
				List<WebElement> pageNumbers = driver.findElements(By.xpath("//*[@id=\"msse\"]/div/form/ul/li"));
				boolean isFound = false;
				for(WebElement currentNumber: pageNumbers) {
					String title = currentNumber.getAttribute("class");
					if(title.equals("active")) {
						isFound=true;
						continue;
					}
					if(isFound) {
						WebElement nextPage = currentNumber.findElement(By.tagName("a"));
						clickAction(nextPage);
						break;
					}
				}
				if(!isFound) {
					break;
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	public boolean processPage() {
		try {
		List<WebElement> links = driver.findElements(By.xpath("//*[@id=\"msse\"]/div/form/div[3]/table/tbody/tr"));
		for(WebElement link: links) {
			String url = link.findElement(By.tagName("a")).getAttribute("href");
			datalist.add(new String[] {url});
		}
		}catch (Exception e) {
			e.printStackTrace();
			return false;
		}
		return true;
	}


	public void processExcel() {
		try {
			FileInputStream inputFile = new FileInputStream(new File("D:\\Freelancer\\thoan_excel\\Results\\" + fileName));
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
						
						processWeb(link, i+1);				
						if((i+1)%50==0) {
							appendExcelFile(fileResult, "");
							System.out.println("===============Write row:" + (i+1));
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

	public void processWeb(String link, int rowIndex) {		
		try {
			driver.get(link);
			Thread.sleep(1000);
			String ISIN = driver.findElement(By.xpath("//*[@id=\"msse\"]/div/div[4]/div[1]/div[3]/table[1]/tbody/tr[2]/td[3]")).getText();
			
			String[] data = new String[] {ISIN};
			datalist.add(data);
			
		} catch (Exception e) {
			e.printStackTrace();
			System.out.println("ERRRRRR: " + rowIndex);

		}
	}
	
	public void clickAction(WebElement element) {
		JavascriptExecutor js = (JavascriptExecutor) driver;
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
