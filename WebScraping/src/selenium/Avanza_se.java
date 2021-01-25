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
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;

public class Avanza_se {
	static WebDriver driver;
	List<String[]> datalist = new ArrayList<>();
	String fileName = "avanza.se.xlsx";
	String fileResult="avanza.se_result.xlsx";

	public Avanza_se() {
		System.setProperty("webdriver.chrome.driver", "D:\\Freelancer\\thoan_excel\\driver79\\chromedriver.exe");
		driver = new ChromeDriver();
		driver.manage().window().maximize();
	}

	public static void main(String[] args) {

		String url = "https://www.avanza.se/aktier/lista.html";
		Avanza_se t = new Avanza_se();
//		t.processMainWeb(url);
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
						processDetailWeb(link, i);
						if(i%50==0) {
							System.out.println("i="+ i);
							appendExcelFile(fileResult, "");
						}
						
					} catch (Exception e) {
						System.out.println("EEEEEERRRRRRRRRR: " + i);
//						e.printStackTrace();
					}
				}
				appendExcelFile(fileResult, "");
				System.out.println("===============Finish====================");
			}
		} catch (Exception e) {
			e.printStackTrace();
		}

	}
	
	public boolean processMainWeb(String url) {
		int pageNumber = 0;
		WebElement totalE = driver.findElement(By.xpath("//*[@id=\"main\"]/div/div/div[5]/div/div[3]/label/span"));
		int total = Integer.parseInt(totalE.getText().split(" ")[0]);
		try {
		while(true) {
			pageNumber++;
			if(pageNumber%5==0) {
				Thread.sleep(10000);
			}
			List<WebElement> links = driver.findElements(By.cssSelector("a[class='ellipsis']"));
			if(links.size()==total) {
				break;
			}
			
			List<WebElement> nextPageEs= driver.findElements(By.xpath("//*[@id=\"main\"]/div/div/div[5]/div/div[5]/div[2]/button"));
			if(!nextPageEs.isEmpty()) {
				clickAction(nextPageEs.get(0));
				Thread.sleep(5000+pageNumber*200);
				System.out.println("Loading..."+pageNumber);
				scrollPage(4000);
			}else {
				break;
			}
		}
		Thread.sleep(10000);
		List<WebElement> links = driver.findElements(By.cssSelector("a[class='ellipsis']"));
		for(WebElement link: links) {
			String a = link.getAttribute("href");
			datalist.add(new String[] {a});
		}
		

		System.out.println("Loading..."+pageNumber);
		}catch (Exception e) {
			e.printStackTrace();
			return false;
		}
		return true;
	}
	
	public void processFrame() {
		
	}
	

	public boolean processDetailWeb(String link, int rowIndex) {
		String[] result= new String[2];
		
		try {
			driver.get(link);
			Thread.sleep(2000);
			String k = driver.findElement(By.xpath("//*[@id=\"surface\"]/div[5]/div[1]/div[3]/div/div[2]/div/div[1]/dl/dd[1]/span")).getText();
			String i = driver.findElement(By.xpath("//*[@id=\"surface\"]/div[5]/div[1]/div[3]/div/div[2]/div/div[1]/dl/dd[2]/span")).getText();
			datalist.add(new String[] {link,i,k});
			
		} catch (Exception e) {
			e.printStackTrace();
//			System.out.println("ERRRRRR: " + rowIndex);
			datalist.add(new String[] {link,"",""});
			return false;
		}
		return true;
	}
	
	public void scrollPage(int length) {
		JavascriptExecutor jse = (JavascriptExecutor)driver;
		jse.executeScript("window.scrollBy(0,"+length+")");
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
}
