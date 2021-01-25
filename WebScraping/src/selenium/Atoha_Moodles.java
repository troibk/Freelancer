package selenium;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Set;

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

public class Atoha_Moodles {
	static WebDriver driver;
	List<String[]> datalist = new ArrayList<>();
	String fileName = "atoha_moodels";
	String fileSheet = "Cost Management";
	String fileResult = "atoha_moodels_result";

	public Atoha_Moodles() {
		System.setProperty("webdriver.chrome.driver", "D:\\Freelancer\\thoan_excel\\driver83\\chromedriver.exe");
		driver = new ChromeDriver();
		driver.manage().window().maximize();
	}

	public static void main(String[] args) {
		int start = 7;
		int end = start;
		Atoha_Moodles t = new Atoha_Moodles();
		t.login();
		t.processWeb(start, end);
		// t.processExcel();

	}

	public void login() {
		String url = "https://atoha.moodle.school/login/index.php";
		try {
			driver.get(url);
			Thread.sleep(2000);
			WebElement username = driver.findElement(By.id("username"));
			username.sendKeys("ha.vu");
			WebElement password = driver.findElement(By.id("password"));
			password.sendKeys("Tlvctg2@");

			WebElement loginBtn = driver.findElement(By.xpath("//*[@id=\"loginbtn\"]"));
			clickAction(loginBtn);

			Thread.sleep(2000);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void processWeb(int begin, int end) {
		try {
			String url = "https://atoha.moodle.school/course/view.php?id=3";
			driver.get(url);
			Thread.sleep(2000);
			List<String[]> links = new ArrayList<>();
			List<WebElement> sections = driver.findElements(By.xpath("//*[@id=\"region-main\"]/div/div/ul/li"));
			for (int i = begin; i <= end; i++) {
				List<WebElement> as = sections.get(i).findElements(By.tagName("a"));
				for (int j = 1; j < as.size(); j++) {
					String link = as.get(j).getAttribute("href");
					String name = as.get(j).findElement(By.tagName("span")).getText().replace("\nĐề thi", "");
					links.add(new String[] {name, link});
				}
			}
			for (String[] link : links) {
				processSubWindow(link[0], link[1]);
			}

		} catch (Exception e) {
			e.printStackTrace();
		}
		System.out.println("===============Finish====================");
	}

	public void processSubWindow(String name, String link) {
		try {
			driver.get(link);
			Thread.sleep(2000);
			String base = driver.getWindowHandle();
			List<WebElement> singlebuttons = driver.findElements(By.cssSelector("div[class='singlebutton']"));
			if (!singlebuttons.isEmpty()) {
				WebElement form = singlebuttons.get(0).findElement(By.tagName("form"))
						.findElement(By.tagName("button"));
				clickAction(form);

				Thread.sleep(2000);

				Set<String> allWindowHandles = driver.getWindowHandles();
				int i = 0;

				for (String handle : allWindowHandles) {
					if (i == 0) {
						i++;
						continue;
					}
					driver.switchTo().window(handle);
					String url = driver.getCurrentUrl();
					System.out.println(url);
					processSubWeb(name);

					driver.close();

				}
				driver.switchTo().window(base);
				Thread.sleep(2000);
			} else {
				List<WebElement> as = driver
						.findElements(By.xpath("//*[@id=\"region-main\"]/div[1]/table/tbody/tr/td[4]/a"));
				if (!as.isEmpty()) {
					clickAction(as.get(0));
					Thread.sleep(2000);
					processSubWeb(name);
				}
			}

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void processSubWeb(String name) {
		try {

			List<WebElement> questions = driver
					.findElements(By.cssSelector("div[class='que multichoice deferredfeedback correct']"));
			if (questions.size() < 2) {
				List<WebElement> incorrects = driver.findElements(By.cssSelector("a[class='qnbutton incorrect free btn btn-secondary']"));
				incorrects.addAll(driver.findElements(By.cssSelector("a[class='qnbutton incorrect free btn btn-secondary flagged']")));
				incorrects.addAll(driver.findElements(By.cssSelector("a[class='qnbutton correct free btn btn-secondary flagged']")));
				List<String> ids = new ArrayList<>();
				for(WebElement in: incorrects) {
					String id = in.getAttribute("id");
					ids.add(id);
				}
				
				for(String id : ids) {
					WebElement btn = driver.findElement(By.id(id));
					clickAction(btn);
					Thread.sleep(500);
					WebElement qce = driver.findElement(By.cssSelector("div[class='formulation clearfix']"));
					WebElement ace = driver.findElement(By.cssSelector("div[class='outcome clearfix']"));
					WebElement qne= driver.findElement(By.cssSelector("span[class='qno']"));
					String qno= qne.getText();
					String qc = qce.getText().replace("Đoạn văn câu hỏi\n", "");
					String ac = ace.getText().replace("Phản hồi\n", "");
					datalist.add(new String[] {name, qno, qc, ac });
				}

			}
			else {
				for (WebElement q : questions) {
					WebElement info = q.findElement(By.cssSelector("div[class='info']"));
					WebElement flag = info.findElement(By.cssSelector("span[class='questionflagtext']"));
					if (flag.getText().equals("Xóa cờ")) {
						WebElement content = q.findElement(By.cssSelector("div[class='content']"));
						WebElement qce = content.findElement(By.cssSelector("div[class='formulation clearfix']"));
						WebElement ace = content.findElement(By.cssSelector("div[class='outcome clearfix']"));
						WebElement qne= info.findElement(By.cssSelector("span[class='qno']"));
						String qno= qne.getText();
						String qc = qce.getText().replace("Đoạn văn câu hỏi\n", "");
						String ac = ace.getText().replace("Phản hồi\n", "");
						datalist.add(new String[] {name, qno, qc, ac });
					}
				}
				
				questions = driver
						.findElements(By.cssSelector("div[class='que multichoice deferredfeedback incorrect']"));
				
				for (WebElement q : questions) {
					WebElement info = q.findElement(By.cssSelector("div[class='info']"));
					WebElement content = q.findElement(By.cssSelector("div[class='content']"));
					WebElement qce = content.findElement(By.cssSelector("div[class='formulation clearfix']"));
					WebElement ace = content.findElement(By.cssSelector("div[class='outcome clearfix']"));
					WebElement qne= info.findElement(By.cssSelector("span[class='qno']"));
					String qno= qne.getText();
					String qc = qce.getText().replace("Đoạn văn câu hỏi\n", "");
					String ac = ace.getText().replace("Phản hồi\n", "");
					datalist.add(new String[] {name, qno, qc, ac });
				}
			}

			System.out.println(name + ":"+datalist.size());
			appendExcelFile(fileName, fileSheet);

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void clickAction(WebElement element) {
		JavascriptExecutor js = (JavascriptExecutor) driver;
		js.executeScript("arguments[0].click();", element);
	}

	public void backAction() {
		JavascriptExecutor js = (JavascriptExecutor) driver;
		js.executeScript("window.history.go(-1)");
	}

	public void selectComboValue(final String elementName, final String value) {
		final Select selectBox = new Select(driver.findElement(By.cssSelector(elementName)));
		selectBox.selectByValue(value);
	}

	private void appendExcelFile(String fileName, String sheetName) {
		Workbook workbook = null;
		Sheet sheet;
		try {
			File file = new File("D:\\Freelancer\\thoan_excel\\Results\\" + fileName + ".xlsx");
			if (!file.exists()) {
				workbook = new XSSFWorkbook();
			} else {
				FileInputStream fip = new FileInputStream(file);
				workbook = new XSSFWorkbook(fip);
			}

			sheet = workbook.getSheet(fileSheet);
			if (sheet == null) {
				sheet = workbook.createSheet(fileSheet);
			}

			int rowNum = sheet.getPhysicalNumberOfRows();
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
