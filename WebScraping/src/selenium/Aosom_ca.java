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
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;

public class Aosom_ca {
	static WebDriver driver;
	List<String[]> datalist = new ArrayList<>();
	List<String[]> datalist2 = new ArrayList<>();

	public Aosom_ca() {
		System.setProperty("webdriver.chrome.driver", "D:\\Freelancer\\thoan_excel\\chromedriver\\chromedriver.exe");
		driver = new ChromeDriver();
		driver.manage().window().maximize();
	}

	public static void main(String[] args) {

		String url = "https://www.aosom.ca/category/patio-furniture~5/?dir=desc&order=bestseller&product_list_limit=100";
		String path ="D:\\Freelancer\\thoan_excel\\Results\\aosom.xlsx";
		Aosom_ca t = new Aosom_ca();
//		t.processWeb(url);
		t.processExcel(path);
	}

	public void processExcel(String path) {
		List<String[]> urls = new ArrayList<>();
		try {
			FileInputStream inputFile = new FileInputStream(new File(path));
			XSSFWorkbook wb = new XSSFWorkbook(inputFile);
			XSSFSheet table1Sheet = wb.getSheetAt(0);

			if (table1Sheet == null) {
				System.out.println("KO CO SHEET :" + table1Sheet);
			} else {

				for (int i = 1; i < table1Sheet.getPhysicalNumberOfRows(); i++) {

					String link = table1Sheet.getRow(i).getCell(0).getStringCellValue();
					urls.add(new String[] { link });

				}

			}

			for (int i = 2; i < urls.size(); i++) {
				boolean result = processFrame(urls.get(i)[0],i + 1);
				if(!result) {
					System.out.println("===============Write Row===============:" + i);
					break;
				}
				if (i % 50 == 0) {
					appendExcelFile("aosom_result");
					System.out.println("===============Write Row===============:" + (i + 1));
				}
			}

			appendExcelFile("aosom_result");
			System.out.println("===============Finish===============");
			inputFile.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void processWeb(String url) {
		driver.get(url);
		try {
			Thread.sleep(10000);
			List<WebElement> items = driver.findElements(By.cssSelector("div[class='product-info-main']"));
			List<String[]> urls = new ArrayList<>();
			for (WebElement item : items) {
				String link = item.findElement(By.tagName("a")).getAttribute("href");
				urls.add(new String[] { link });
			}
			String lastPage = driver.getCurrentUrl();
			while (true) {

				System.out.println(urls.size());

				List<WebElement> pagingElement = driver.findElements(
						By.xpath("//*[@id=\"app\"]/div[4]/main/div/div/div/div[2]/div[2]/div/div[2]/div[4]/ul/li"));
				WebElement pagingNextBtn = pagingElement.get(pagingElement.size() - 1)
						.findElement(By.tagName("button"));
				JavascriptExecutor js = ((JavascriptExecutor) driver);
				js.executeScript("arguments[0].click();", pagingNextBtn);
				Thread.sleep(10000);
				String currentUrl = driver.getCurrentUrl();
				if (lastPage.equals(currentUrl)) {
					break;
				}
				lastPage = currentUrl;
				items = driver.findElements(By.cssSelector("div[class='product-info-main']"));
				for (WebElement item : items) {
					String link = item.findElement(By.tagName("a")).getAttribute("href");
					urls.add(new String[] { link });
				}

			}

			writeExcelFile("aosom", urls);

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public boolean processFrame(String url, int rowIndex) {
		driver.get(url);
		try {
			Thread.sleep(2000);
			String name = driver.findElement(By.xpath("//*[@id=\"right-product-info\"]/h1")).getText();
			String imageUrl = driver.findElement(By.xpath("//*[@id=\"product-zoomer-gallery\"]/div[1]/div/img")).getAttribute("src");
			List<WebElement> priceList = driver.findElements(By.xpath("//*[@id=\"right-product-info\"]/div[2]/div[1]/span"));
			String salePrice = "";
			String price="";
			if(priceList.size()==1) {
				price=priceList.get(0).getText();
				salePrice=price;
			}else if(priceList.size()==2) {
				price=priceList.get(0).getText();
				salePrice = priceList.get(1).getText();
			}
			String description = driver.findElements(By.cssSelector("div[class='v-window-item']")).get(0).findElement(By.tagName("article")).getText();
			String reviews = driver.findElement(By.xpath("//*[@id=\"comment-container\"]/div[1]/div[1]/div[1]/div[2]")).getText();
			String sku = driver.findElement(By.xpath("//*[@id=\"right-product-info\"]/p[2]")).getText();
			
			String[] data = new String[7];
			data[0]=name;
			data[1]=imageUrl;
			data[2]=salePrice;
			data[3]=price;
			data[4]=description;
			data[5]=reviews;
			data[6]=sku;
			datalist.add(data);
			
		} catch (Exception e) {
			e.printStackTrace();
			System.out.println("EEEEEEEE:" + rowIndex);
			return false;
		}
		return true;
	}

	public void selectComboValue(final String elementName, final String value) {
		final Select selectBox = new Select(driver.findElement(By.name(elementName)));
		selectBox.selectByValue(value);
	}

	private void appendExcelFile(String fileName) {
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
			int rowNum = 0;
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
		Workbook workbook = new XSSFWorkbook(); // new HSSFWorkbook() for generating `.xls` file
		CreationHelper createHelper = workbook.getCreationHelper();

		// Create a Sheet
		Sheet sheet = workbook.createSheet("Table");

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

	private void writeExcelFile(String fileName, List<String[]> data) {
		Workbook workbook = new XSSFWorkbook(); // new HSSFWorkbook() for generating `.xls` file
		CreationHelper createHelper = workbook.getCreationHelper();

		// Create a Sheet
		Sheet sheet = workbook.createSheet("Table");

		CellStyle dateCellStyle = workbook.createCellStyle();
		dateCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd-MM-yyyy"));

		// Create Other rows and cells with employees data
		int rowNum = 1;
		for (String[] d : data) {
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
