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

public class Vivino_com {
	static WebDriver driver;
	List<String[]> datalist = new ArrayList<>();
	List<String[]> datalist2 = new ArrayList<>();

	public Vivino_com() {
		System.setProperty("webdriver.chrome.driver", "D:\\Freelancer\\thoan_excel\\chromedriver\\chromedriver.exe");
		driver = new ChromeDriver();
		driver.manage().window().maximize();
	}

	public static void main(String[] args) {

		String url = "https://www.vivino.com/explore?e=eJzLLbI1VMvNzLM1NFDLTaywNTUwUEuutA3zU0u2DQ12USsASqen2ZYlFmWmliTmqOUm26rlJwGxbUpqcbJaeUl0LFAFmDICAIccF_w=";

		Vivino_com t = new Vivino_com();
		t.processWeb(url);

	}

	public void processWeb(String url) {
		// driver.get(url);
		try {
			// Thread.sleep(5000);
			// String totalItemsS =
			// driver.findElement(By.xpath("//*[@id=\"explore-page-app\"]/div/div/h2")).getText().split("\\s+")[1];
			// int totalItems = Integer.parseInt(totalItemsS);
			// List<String[]> urls = new ArrayList<>();
			// while(true) {
			// JavascriptExecutor js = ((JavascriptExecutor) driver);
			// js.executeScript("window.scrollTo(0, document.body.scrollHeight)");
			// Thread.sleep(5000);
			// List<WebElement> items =
			// driver.findElements(By.cssSelector("div[class='explorerCard__explorerCard--3Q7_0
			// explorerPageResults__explorerCard--3q6Qe']"));
			// System.out.println(items.size());
			// if(items.size()>=totalItems) {
			// for(WebElement item: items) {
			// WebElement element =
			// item.findElement(By.cssSelector("div[class='vintageTitle__vintageTitle--2iCdc']"));
			// String link = element.findElement(By.tagName("a")).getAttribute("href");
			// String[] data = new String[1];
			// data[0]=link;
			// urls.add(data);
			// }
			// break;
			// }
			// }
			//
			// writeExcelFile("vivino", urls);
			List<String[]> urls = new ArrayList<>();
			FileInputStream inputFile = new FileInputStream(new File("D:\\Freelancer\\thoan_excel\\Results\\vivino.xlsx"));
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

			inputFile.close();

			for (String[] link : urls) {
				try {
					String id = link[0].split("\\?")[0].substring(link[0].lastIndexOf("/") + 1);
					driver.get(link[0]);
					Thread.sleep(1000);

					String bottleName = driver.findElement(By.cssSelector("h1[class='winePageHeader__vintage--2Vux3']"))
							.getText();
					String year = bottleName.substring(bottleName.lastIndexOf(' ') + 1);
					WebElement wineLocation = driver.findElement(By.xpath(
							"//*[@id=\"vintage-page-app\"]/div[1]/div/div/div[1]/div/div[2]/div[1]/div/div/div[2]"));
					String type = wineLocation.getText();
					List<WebElement> wineRegion = wineLocation.findElements(By.tagName("a"));
					String country = wineRegion.get(1).getText();
					List<WebElement> summaryList = driver
							.findElements(By.cssSelector("div[class='wineSummary__fact--9Af9b']"));
					String winery = summaryList.get(0)
							.findElement(By.cssSelector("div[class='wineSummary__factValue--kegJ-']")).getText();
					String foodPairing = summaryList.get(4)
							.findElement(By.cssSelector("div[class='wineSummary__factValue--kegJ-']")).getText();
					String grapes = summaryList.get(1)
							.findElement(By.cssSelector("div[class='wineSummary__factValue--kegJ-']")).getText();
					String regionalStyles = summaryList.get(3)
							.findElement(By.cssSelector("div[class='wineSummary__factValue--kegJ-']")).getText();
					String region = summaryList.get(2)
							.findElement(By.cssSelector("div[class='wineSummary__factValue--kegJ-']")).getText();
					String priceMsg = driver
							.findElement(By.cssSelector("div[class='purchaseAvailabilityPPC__notSoldMessage--3t8VH']"))
							.getText();
					String price = priceMsg.substring(priceMsg.lastIndexOf(' '));
					String rating = driver.findElement(By.cssSelector("span[class='vivinoRating__rating--4Oti3']"))
							.getText();
					String countOfRatings = driver
							.findElement(By.cssSelector("span[class='vivinoRating__ratingCount--NmiVg']")).getText();
					String alcoholPercent = summaryList.get(5)
							.findElement(By.cssSelector("div[class='wineSummary__factValue--kegJ-']")).getText();
					String[] data = new String[15];
					data[0] = winery;
					data[1] = bottleName;
					data[2] = id;
					data[3] = year;
					data[4] = type;
					data[5] = region;
					data[6] = foodPairing;
					data[7] = grapes;
					data[8] = regionalStyles;
					data[9] = country;
					data[10] = price;
					data[11] = rating;
					data[12] = countOfRatings;
					data[13] = alcoholPercent;
					data[14] = link[0];
					datalist.add(data);

					List<WebElement> allReviews = driver
							.findElements(By.xpath("//*[@id=\"all_reviews\"]/div/div[2]/div/div"));

					while (true) {
						try {
							WebElement showMore = driver
									.findElement(By.xpath("//*[@id=\"all_reviews\"]/div/div[2]/div/a"));
							showMore.click();
							Thread.sleep(2000);
							List<WebElement> tmp = driver
									.findElements(By.xpath("//*[@id=\"all_reviews\"]/div/div[2]/div/div"));
							if (allReviews.size() == tmp.size()) {
								allReviews = tmp;
								break;
							}
							allReviews = tmp;
						} catch (Exception e) {
							e.printStackTrace();
							break;
						}
					}

					for (WebElement review : allReviews) {
						String ratedDate = review.findElement(By.cssSelector("a[class='anchor__anchor--2QZvA communityReviewer__ratedOn--yAWY6']")).getText();
						WebElement ratingE = review.findElement(By.cssSelector("div[class='rating__rating--ZZb_x rating__user--15hMB']"));
						List<WebElement> icon100 = ratingE.findElements(By.cssSelector("i[class='rating__icon--2T9_0 rating__icon100--2vw_3']"));
						List<WebElement> icon50 = ratingE.findElements(By.cssSelector("i[class='rating__icon--2T9_0 rating__icon50--3xGES']"));
						String stars;
						if (icon50.isEmpty()) {
							stars = "" + icon100.size();
						} else {
							stars = "" + icon100.size() + "," + 5;
						}

						String name = review.findElement(By.cssSelector("a[class='anchor__anchor--2QZvA communityReviewer__alias--3JFXY']")).getText();
						String userID = review.findElement(By.cssSelector("a[class='anchor__anchor--2QZvA communityReviewer__alias--3JFXY']")).getAttribute("href");
						String reviewContent = review.findElement(By.cssSelector("div[class='communityReview__textSection--vu-i-']")).getText();

						String[] reviewData = new String[7];
						reviewData[0] = id;
						reviewData[1] = year;
						reviewData[2] = ratedDate;
						reviewData[3] = stars;
						reviewData[4] = name;
						reviewData[5] = userID;
						reviewData[6] = reviewContent;
						datalist2.add(reviewData);
					}
				} catch (Exception e) {
					System.out.println(link[0]);
					e.printStackTrace();
				}

			}
			writeExcelFile("vivino_result",datalist);
			writeExcelFile("vivino_review",datalist2);
		} catch (Exception e) {
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

			if (workbook.getNumberOfSheets() == 0) {
				sheet = workbook.createSheet("Results");
			} else {
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
