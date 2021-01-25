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

public class IntBar_org {
	static WebDriver driver;
	List<String[]> datalist = new ArrayList<>();
	String fileName = "int-bar.org.xlsx";
	String fileResult="int-bar.org_result.xlsx";

	public IntBar_org() {
		System.setProperty("webdriver.chrome.driver", "D:\\Freelancer\\thoan_excel\\driver\\chromedriver.exe");
		driver = new ChromeDriver();
		driver.manage().window().maximize();
	}

	public static void main(String[] args) {

		String url = "https://www.int-bar.org/MemberDirectory/Search.cfm";
		IntBar_org t = new IntBar_org();
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
		
		doLogin();
		try {
			driver.get(url);
			Thread.sleep(3000);
			String[] countries = new String[] {"147","149","153","155","158","159"};
			for (String c : countries) {
				Select selectBox = new Select(driver.findElement(By.id("country")));
				selectBox.selectByValue(c);
				WebElement searchBtn = driver
						.findElement(By.xpath("//*[@id=\"mainForm\"]/div/div[3]/div/div/div/button"));
				searchBtn.click();
				Thread.sleep(3000);
				while (true) {
					int pageSize = processPage();
					if (pageSize > 0) {
						String currentPage = driver.findElement(By.xpath("//*[@id=\"mainForm\"]/div[2]/div/div[2]"))
								.getText();
						System.out.println("===========Current Page: " + currentPage + ", Size: " + datalist.size());

						List<WebElement> pagingBtn = driver
								.findElements(By.xpath("//*[@id=\"mainForm\"]/div[2]/div/div[3]/a"));
						boolean isNext = false;
						for (WebElement p : pagingBtn) {
							if (p.getAttribute("title").equals("Next")) {
								p.click();
								isNext = true;
								break;
							}
						}
						if (!isNext) {
							break;
						}
						Thread.sleep(3000);
					}else {
						break;
					}
				}
				appendExcelFile(fileName, "");
				System.out.println("============Write Country: " + c);
				WebElement searchAgainBtn = driver
						.findElement(By.xpath("//*[@id=\"mainForm\"]/div[1]/div/div[2]/div/button[1]"));
				searchAgainBtn.click();
				Thread.sleep(5000);
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		System.out.println("===============Finish====================");
	}

	public int processPage() {
		List<WebElement> rows = driver.findElements(By.xpath("//*[@id=\"mainForm\"]/div[3]/div/table/tbody/tr"));
		for (int i =1; i<rows.size();i++) {
			List<WebElement> tds = rows.get(i).findElements(By.tagName("td"));
			String name = tds.get(1).getText();
			String link = tds.get(1).findElement(By.tagName("a")).getAttribute("href");
			String firm = tds.get(2).getText();
			String country = tds.get(3).getText();
			String[] data = new String[4];
			data[0] = link;
			data[1]=name;
			data[2] = firm;
			data[3] = country;
			datalist.add(data);
		}
		return rows.size();
	}

	public void processExcel() {
		doLogin();
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
						String name = table1Sheet.getRow(i).getCell(1).getStringCellValue();
						String firm = table1Sheet.getRow(i).getCell(2).getStringCellValue();
						String country= table1Sheet.getRow(i).getCell(3).getStringCellValue();
						String[] email_state = processWeb(link, i + 1);
						if(email_state==null) {
							break;
						}
						String[] data= new String[4];
						data[0]=firm;
						data[1]=name;
						data[2]=country;
						data[3]=email_state[0];
//						data[4]=email_state[1];
						datalist.add(data);
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

	public String[] processWeb(String link, int rowIndex) {
		String[] result= new String[2];
		
		try {
			driver.get(link);
			Thread.sleep(2000);
			String email="";
			List<WebElement> emaislE = driver.findElements(By.id("ctl00_MainContent_lnkEmailAddress"));
			if(!emaislE.isEmpty()) {
				email= emaislE.get(0).getText();
			}
			
//			WebElement addressE = driver.findElement(By.xpath("//*[@id=\"ctl00_MainContentContainer\"]/div/div/div[1]/div[1]/div[2]"));
//			String address = addressE.getText();
			result[0]=email;
//			result[1]=address;
			
		} catch (Exception e) {
			e.printStackTrace();
			System.out.println("ERRRRRR: " + rowIndex);
			return null;
		}
		return result;
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
