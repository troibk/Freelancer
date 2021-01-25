package htmlparser;
import java.awt.Dimension;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.filechooser.FileNameExtensionFilter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import thoan_excel.ExcelProcessing;
import thoan_excel.FirmData;
import thoan_excel.Model;
 


public class HtmlParser extends JPanel implements ActionListener{
	private static final long serialVersionUID = 1L;

	JButton go;
	JFileChooser fileChoosen;
	String choosertitle;
	File myFile;
    private static List<String[]> datalist = new ArrayList<>();
    private static List<Email> emailList = new ArrayList<>();
	public HtmlParser() {
		go = new JButton("Do it");
		go.addActionListener(this);
		add(go);
	}

	public void actionPerformed(ActionEvent e) {
		fileChoosen = new JFileChooser();
		fileChoosen.setCurrentDirectory(new java.io.File("."));
		fileChoosen.setDialogTitle(choosertitle);
		fileChoosen.setFileSelectionMode(JFileChooser.FILES_ONLY);
		//
		// disable the "All files" option.
		//
		fileChoosen.setAcceptAllFileFilterUsed(false);
		fileChoosen.addChoosableFileFilter(new FileNameExtensionFilter("*.xml", "xml"));
		//
		if (fileChoosen.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
			System.out.println("getCurrentDirectory(): " + fileChoosen.getCurrentDirectory());
			System.out.println("getSelectedFile() : " + fileChoosen.getSelectedFile());
			myFile = fileChoosen.getSelectedFile();
			processXML(myFile);
		} else {
			System.out.println("No Selection ");
		}

	}
	
	

	public Dimension getPreferredSize() {
		return new Dimension(100, 100);
	}

	public void processXML(File file) {
		
		Document doc = null;
		try {
			
			   // String HTMLSTring = readFileAsString(file);
			 
			 doc = Jsoup.parse(file, "UTF-8");
			 Elements items = doc.select("us-patent-application");
			 
			 for (Element element: items) {
				 Element card = element.selectFirst("document-id");
				 Element header = card.selectFirst("div.card__header");
				 String name = header.text();
				 String link= card.attr("href");
				 datalist.add(new String[] {name,link});
			}
				
				
			

		} catch (Exception ioe) {
			ioe.printStackTrace();
		}

		
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

	public static void main(String s[]) {
		JFrame frame = new JFrame("");
		HtmlParser panel = new HtmlParser();
		frame.addWindowListener(new WindowAdapter() {
			public void windowClosing(WindowEvent e) {
				System.exit(0);
			}
		});
		frame.getContentPane().add(panel, "Center");
		frame.setSize(panel.getPreferredSize());
		frame.setVisible(true);
	}

    
    public String readFileAsString(File file) {
        String text = "";
        try {
          text = new String(Files.readAllBytes(file.toPath()));
        } catch (IOException e) {
          e.printStackTrace();
        }

        return text;
      }
 
}