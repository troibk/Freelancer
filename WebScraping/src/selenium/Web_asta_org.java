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

public class Web_asta_org {
	static WebDriver driver;
	List<String[]> datalist = new ArrayList<>();
	String fileName = "web.asta.org.xlsx";
	String fileResult = "web.asta.org_result.xlsx";
	String categories="inch seal sizes";

	public Web_asta_org() {
		System.setProperty("webdriver.chrome.driver", "D:\\Freelancer\\thoan_excel\\chromedriver\\chromedriver.exe");
		driver = new ChromeDriver();
		driver.manage().window().maximize();
	}

	public static void main(String[] args) {

		String url = "https://web.asta.org/iMIS/ASTA/Contacts/People_Search.aspx?WebsiteKey=abd5fcd9-388c-4205-80a4-9fe223227405&navItemNumber=11304";
		Web_asta_org t = new Web_asta_org();
//		t.processWeb(url);
		t.processExcel();

	}


	public void processWeb(String url) {

		try {
			driver.get(url);
			Thread.sleep(3000);
			
			WebElement loginBtn = driver.findElement(By.xpath("//*[@id=\"ctl01_TemplateBody_WebPartManager1_gwpciBPDirectorySearch_ciBPDirectorySearch_sbtnSearch\"]"));
			clickAction(loginBtn);
				while(true) {
					if(processPage()) {
						appendExcelFile(fileName, "");
						
					
					List<WebElement> items = driver.findElements(By.xpath("//*[@id=\"ctl01_TemplateBody_WebPartManager1_gwpciBPDirectorySearch_ciBPDirectorySearch_gvResults\"]/tbody/tr[1]/td/table/tbody/tr/td"));
					for(int i =0; i<items.size();i++) {
						List<WebElement> spans = items.get(i).findElements(By.tagName("span"));
						if(!spans.isEmpty()) {
							WebElement nextPage = items.get(i+1);
							clickAction(nextPage.findElement(By.tagName("a")));
							break;
						}
					}
					}else {
						break;
					}
					
				}
			
		} catch (Exception e) {
			e.printStackTrace();
		}
		/*appendExcelFile(fileName, "");
		System.out.println("===============Finish====================");*/
	}

	public boolean processPage() {
		try {
			
			Thread.sleep(3000);
			List<WebElement> rows = driver.findElements(By.xpath("//*[@id=\"ctl01_TemplateBody_WebPartManager1_gwpciBPDirectorySearch_ciBPDirectorySearch_gvResults\"]/tbody/tr"));
			for(int i = 2; i<rows.size()-2; i++) {
				List<WebElement> tds = rows.get(i).findElements(By.tagName("td"));
				WebElement nameE = tds.get(0).findElement(By.tagName("a"));
				String link = nameE.getAttribute("href");
				String name = nameE.getText();
				String company = tds.get(1).getText();
				String title = tds.get(2).getText();
				String city = tds.get(3).getText();
				String state = tds.get(4).getText();
				String[] data = new String[6];
				data[0]=link;
				data[1]=name;
				data[2]=company;
				data[3]=title;
				data[4]=city;
				data[5]=state;
				datalist.add(data);
			}
			
			
			System.out.println("==========Current Page: "+ rows.get(rows.size()-2).findElement(By.tagName("td")).getText());
			
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}
		return true;
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

				for (int i = 91; i < table1Sheet.getPhysicalNumberOfRows(); i++) {
					try {

						Cell cell = table1Sheet.getRow(i).getCell(0);
						if (cell == null) {
							continue;
						}

						String link = table1Sheet.getRow(i).getCell(0).getStringCellValue();
						String name = table1Sheet.getRow(i).getCell(1).getStringCellValue();
						String company = table1Sheet.getRow(i).getCell(2).getStringCellValue();
						String title = table1Sheet.getRow(i).getCell(3).getStringCellValue();
						String city=table1Sheet.getRow(i).getCell(4).getStringCellValue();
						String state =table1Sheet.getRow(i).getCell(5).getStringCellValue();
						
						if ((i + 1) % 50 == 0) {
							appendExcelFile(fileResult, "");
							System.out.println("===============Write row:" + (i + 1));
						}
						
						if(!processWeb(link, name, company, title, city, state, (i+1))) {
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

	public boolean processWeb(String link,String name,String company,String title,String city,String state, int rowIndex) {

		try {
			driver.get(link);
			Thread.sleep(3000);
			List<WebElement> addressE =driver.findElements(By.xpath("//*[@id=\"ctl01_TemplateBody_WebPartManager1_gwpciProfile_ciProfile_contactAddress__address\"]")); 
			String address = addressE.isEmpty() ? "": addressE.get(0).getText();
			List<WebElement> phoneE =driver.findElements(By.xpath("//*[@id=\"ctl01_TemplateBody_WebPartManager1_gwpciProfile_ciProfile_contactAddress__phoneNumber\"]")); 
			String phone = phoneE.isEmpty() ? "": phoneE.get(0).getText();
			List<WebElement> emailE =driver.findElements(By.xpath("//*[@id=\"ctl01_TemplateBody_WebPartManager1_gwpciProfile_ciProfile_contactAddress__email\"]")); 
			String email = emailE.isEmpty() ? "": emailE.get(0).getText();
			List<WebElement> webE =driver.findElements(By.xpath("//*[@id=\"ctl01_TemplateBody_WebPartManager1_gwpciProfileSection_ciProfileSection_CsContact.Website\"]")); 
			String web = webE.isEmpty() ? "": webE.get(0).getText();
			
			String aboutMe = driver.findElement(By.xpath("//*[@id=\"ctl01_TemplateBody_WebPartManager1_gwpciAreasofExpertise_ciAreasofExpertise_singleInstancePanel\"]")).getText();
			
			String[] data = new String[10];
			data[0]=name;
			data[1]=company;
			data[2]=title;
			data[3]=city;
			data[4]=state;
			data[5]=address;
			data[6]=phone;
			data[7]=email;
			data[8]=web;
			data[9]=aboutMe;
			datalist.add(data);
			
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
