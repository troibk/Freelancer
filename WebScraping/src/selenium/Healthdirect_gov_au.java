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

public class Healthdirect_gov_au {
	static WebDriver driver;
	List<String[]> datalist = new ArrayList<>();
	String fileName = "healthdirect.gov.au.xlsx";
	String fileResult="healthdirect.gov.au_result.xlsx";

	public Healthdirect_gov_au() {
		System.setProperty("webdriver.chrome.driver", "D:\\Freelancer\\thoan_excel\\driver79\\chromedriver.exe");
		driver = new ChromeDriver();
		driver.manage().window().maximize();
	}

	public static void main(String[] args) {

		String url = "https://www.healthdirect.gov.au/australian-health-services/results/the_rocks-2000/tihcs-aht-11222/gp-general-practice?pageIndex=1&tab=SITE_VISIT";
		String key="the_rocks-2000";
		Healthdirect_gov_au t = new Healthdirect_gov_au();
//		t.processMainWeb(url, key);
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
						String key= table1Sheet.getRow(i).getCell(1).getStringCellValue();
						String link = table1Sheet.getRow(i).getCell(1).getStringCellValue();
						processDetailWeb(key, link, i);
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
	
	public boolean processMainWeb(String url, String key) {

		try {
		driver.get(url);
		Thread.sleep(2000);
		
		processFrame(key);
		
//		while(true) {
//			String total = driver.findElement(By.xpath("//*[@id=\"paginatinLabel\"]")).getText();
//			
//			if(processFrame(key)) {
//				System.out.println("Loading..."+total);
//			}
//			
//			WebElement nextPageE= driver.findElement(By.xpath("//*[@id=\"SITE_VISIT\"]/div[2]/div[11]/nav/a[12]"));
//			if(nextPageE.getAttribute("aria-disabled").equals("true")) {
//				break;
//			}else{
//				clickAction(nextPageE);
//				Thread.sleep(2000);
//			}
//		}
		}catch (Exception e) {
			e.printStackTrace();
			return false;
		}
		
		appendExcelFile(fileName, "");
		System.out.println();
		
		return true;
	}
	
	public boolean processFrame(String key) {
		List<WebElement> allE = driver.findElements(By.xpath("//*[@id=\"SITE_VISIT\"]/div[2]/div"));
		for(WebElement a : allE) {
		List<WebElement> links = a.findElements(By.tagName("a"));
		for(WebElement l : links) {
			String link = l.getAttribute("href");
			datalist.add(new String[] {key,link});
		}
		}
		return true;
	}
	

	public boolean processDetailWeb(String key, String link, int rowIndex) {
		
		try {
			driver.get(link);
			Thread.sleep(5000);
			
			String name = driver.findElement(By.xpath("/html/body/main/div/header/h1")).getText();
			String address = driver.findElement(By.xpath("/html/body/main/div/section/div[2]/div/div[2]/div/p[1]")).getText();
			
			List<WebElement> contactsE = driver.findElement(By.cssSelector("div[class='veyron-hsf-contact-details']")).findElements(By.tagName("a"));
			String phone = "";
			String email="";
			String web= "";
			for(WebElement c : contactsE) {
				String r = c.getAttribute("href");
				if(r.contains("tel")) {
					phone=r;
				}else if (r.contains("mail")) {
					email=r;
				}else if(r.contains("http")) {
					web=r;
				}
			}
			List<WebElement> opElements = driver.findElements(By.cssSelector("div[class='hsf-service_details-data-hours']"));
			String open = "";
			if(!opElements.isEmpty()) {
				open=opElements.get(0).getText();
			}
			String billing="";
			List<WebElement> billingE=driver.findElements(By.cssSelector("p[class='hsf-service_details-data-billing']"));
			if(!billingE.isEmpty()) {
				billing=billingE.get(0).getText();
			}
			
			String appoinment = "";
			List<WebElement> apE=driver.findElements(By.cssSelector("p[class='hsf-service_details-data-description']"));
			if(!apE.isEmpty()) {
				appoinment=apE.get(0).getText();
			}
			datalist.add(new String[] {key,link,name,address,phone,email, web, open, billing, appoinment});
			
		} catch (Exception e) {
			e.printStackTrace();
//			System.out.println("ERRRRRR: " + rowIndex);
			datalist.add(new String[] {key, link,"",""});
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
