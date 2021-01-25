package selenium;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;

public class Kybar_org {
	static WebDriver driver;
	List<String[]> datalist = new ArrayList<>();

	public Kybar_org() {
		System.setProperty("webdriver.chrome.driver", "D:\\Freelancer\\thoan_excel\\chromedriver\\chromedriver.exe");
		driver = new ChromeDriver();
		driver.manage().window().maximize();
	}

	public static void main(String[] args) {

		String url = "https://www.kybar.org/search/custom.asp?id=2947";

		Kybar_org t = new Kybar_org();
		t.processWeb(url);

	}

	public void processWeb(String url) {
		driver.get(url);
		try {
			Thread.sleep(1000);
			String comboName = "cdlCustomFieldValueIDAreasofPractice-SINGLE";
			WebElement select = driver.findElement(By.name(comboName));
			List<WebElement> options = select.findElements(By.tagName("option"));
			for (WebElement option : options) {
				String comboValue = option.getAttribute("value");
				if(comboValue.isEmpty()) {
					continue;
				}
				selectComboValue(comboName, comboValue);
				WebElement element = driver.findElement(By.xpath("//*[@id=\"main\"]/table/tbody/tr[8]/td[2]/input"));
				JavascriptExecutor js = (JavascriptExecutor) driver;
				js.executeScript("window.scrollTo(0," + element.getLocation().x + ")");
				element.click();

				Thread.sleep(5000);

				// WebElement myDynamicElement = (new WebDriverWait(driver,
				// 30)).until(ExpectedConditions.presenceOfElementLocated(By.xpath("//*[@id=\"SearchResultsFrame\"]")));

				driver.switchTo().frame(driver.findElement(By.id("SearchResultsFrame")));
				processFrame();
				System.out.println(datalist.size());
				int currentPage =9;
				String totalPageS = driver.findElement(By.cssSelector("tr[class='DotNetPager']")).findElements(By.tagName("span")).get(1).getText();
				int totalPage = Integer.parseInt(totalPageS.substring(totalPageS.length()-3).trim());
				while (true) {
					try {
						List<WebElement> nextPages = driver.findElement(By.cssSelector("tr[class='DotNetPager']")).findElements(By.tagName("a"));
						WebElement nextPage=nextPages.get(currentPage);
						WebElement lastPage = nextPages.get(nextPages.size()-1);
						currentPage++;
						if(currentPage==nextPages.size()) {
							currentPage=1;
						}
						String pageNum = driver.findElement(By.cssSelector("tr[class='DotNetPager']")).findElement(By.tagName("span")).getText();
						//TODO
						
						
						System.out.println("XXXXXX"+pageNum);
						String script = nextPage.getAttribute("href");
						js.executeScript(script);
						Thread.sleep(5000);
						processFrame();
						System.out.println(datalist.size());
					} catch (Exception e) {
						e.printStackTrace();
						break;
					}
				}
				appendExcelFile("kyba", comboValue);
				datalist.clear();
			}
		} catch (InterruptedException e) {
			e.printStackTrace();
		}
	}

	public void processFrame() {
		WebElement table = driver.findElement(By.id("SearchResultsGrid"));
		List<WebElement> els = table.findElements(By.xpath("//*[@id=\"SearchResultsGrid\"]/tbody/tr"));

		for (int i = 0; i < els.size(); i++) {
			List<WebElement> divs = els.get(i).findElements(
					By.xpath("//*[@id=\"SearchResultsGrid\"]/tbody/tr[" + (i + 1) + "]/td/table/tbody/tr/td[1]/div"));
			if (divs.size() > 0) {
				String name = divs.get(0).getText();
				String link = divs.get(0).findElement(By.tagName("a")).getAttribute("href");
				String state = "";
				if (divs.size() > 4) {
					state = divs.get(4).getText();
				}

				String[] data = new String[3];
				data[0] = link;
				data[1] = name;
				data[2] = state;
				datalist.add(data);
			}
		}
	}

	public void selectComboValue(final String elementName, final String value) {
		final Select selectBox = new Select(driver.findElement(By.name(elementName)));
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
			
			if(workbook.getNumberOfSheets() == 0) {
				sheet = workbook.createSheet("Results");
			}else {
				sheet = workbook.getSheetAt(0);
			}
			int rowNum = 1;
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

	private void writeExcelFile(String fileName) {
		// String[] columns = {"Company", "Phone", "Website"};
		Workbook workbook = new XSSFWorkbook(); // new HSSFWorkbook() for generating `.xls` file

		/*
		 * CreationHelper helps us create instances of various things like DataFormat,
		 * Hyperlink, RichTextString etc, in a format (HSSF, XSSF) independent way
		 */
		CreationHelper createHelper = workbook.getCreationHelper();

		// Create a Sheet
		Sheet sheet = workbook.createSheet("Table");

		// Create a Font for styling header cells
		// Font headerFont = workbook.createFont();
		// headerFont.setBold(true);
		// headerFont.setFontHeightInPoints((short) 14);
		// headerFont.setColor(IndexedColors.RED.getIndex());
		//
		// // Create a CellStyle with the font
		// CellStyle headerCellStyle = workbook.createCellStyle();
		// headerCellStyle.setFont(headerFont);
		//
		// // Create a Row
		// Row headerRow = sheet.createRow(0);
		//
		// // Create cells
		// for(int i = 0; i < columns.length; i++) {
		// Cell cell = headerRow.createCell(i);
		// cell.setCellValue(columns[i]);
		// cell.setCellStyle(headerCellStyle);
		// }

		// Create Cell Style for formatting Date
		CellStyle dateCellStyle = workbook.createCellStyle();
		dateCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd-MM-yyyy"));

		// Create Other rows and cells with employees data
		int rowNum = 1;
		for (String[] d : datalist) {
			Row row = sheet.createRow(rowNum++);
			for (int i = 0; i < d.length; i++) {
				row.createCell(i).setCellValue(d[i]);
			}
		}

		// Resize all columns to fit the content size
		// for(int i = 0; i < columns.length; i++) {
		// sheet.autoSizeColumn(i);
		// }

		// Write the output to a file
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
