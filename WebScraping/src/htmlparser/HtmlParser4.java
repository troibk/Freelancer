package htmlparser;
import java.awt.Dimension;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.List;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JPanel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
 


public class HtmlParser4 extends JPanel implements ActionListener{
	private static final long serialVersionUID = 1L;

	JButton go;
	JFileChooser chooser;
	String choosertitle;
	File folder;
    private static List<String[]> datalist = new ArrayList<>();

	public HtmlParser4() {
		go = new JButton("Do it");
		go.addActionListener(this);
		add(go);
	}

	public void actionPerformed(ActionEvent e) {
		chooser = new JFileChooser();
		chooser.setCurrentDirectory(new java.io.File("."));
		chooser.setDialogTitle(choosertitle);
		chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
		//
		// disable the "All files" option.
		//
		chooser.setAcceptAllFileFilterUsed(false);
		//
		if (chooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
			System.out.println("getCurrentDirectory(): " + chooser.getCurrentDirectory());
			System.out.println("getSelectedFile() : " + chooser.getSelectedFile());
			folder = chooser.getSelectedFile();
		} else {
			System.out.println("No Selection ");
		}
		List<File> txtFiles = new ArrayList<>();
		if (folder != null && folder.isDirectory()) {
			File[] listOfFiles = folder.listFiles();
			String fileExt = "";
			for (File f : listOfFiles) {
				if (!f.isHidden()) {
					fileExt = f.getName();
					if (fileExt.endsWith(".html") || fileExt.endsWith(".HTML")) {
						txtFiles.add(f);
					}
				}
			}
		}
		int i = 0;
		for (File f : txtFiles) {
			processTXT(f);
			i++;
			
			System.out.println("YYYYY" + datalist.size());
			
		}
		writeExcelFile("projectviewercentral");
		System.out.println("$$$$$$$$$$FINISH$$$$$$$$$$$$$");
	}

	public Dimension getPreferredSize() {
		return new Dimension(100, 100);
	}

	public void processTXT(File file) {
		
		Document doc = null;
		try {
			
			   // String HTMLSTring = readFileAsString(file);
			 
			 doc = Jsoup.parse(file, "UTF-8");
			 Elements apply_infos = doc.select("table.table").first().select("tr");
			 
			 for (Element element: apply_infos) {
				 try {
					 String[] data = new String[2];
					 Element website = element.select("span.pl-results-url").first();
					 String email="";
					 try {
					 Element email_e = element.select("span.pl-results-email").first();
					 email=email_e.select("a").first().attr("href");
					 
					 }catch (Exception e) {
						email="Email";
					}
					 String url=website.select("a").first().attr("href");
					 data[0]=url;
					 data[1]=email;
					 datalist.add(data);
				 }catch (Exception e) {
					
				}
			 }

			

		} catch (Exception ioe) {
			ioe.printStackTrace();
		}

		
	}
	
	
	private void writeExcelFile(String fileName) {
		String[] columns = {"Link", "Email"};
		Workbook workbook = new XSSFWorkbook(); // new HSSFWorkbook() for generating `.xls` file

        /* CreationHelper helps us create instances of various things like DataFormat, 
           Hyperlink, RichTextString etc, in a format (HSSF, XSSF) independent way */
        CreationHelper createHelper = workbook.getCreationHelper();

        // Create a Sheet
        Sheet sheet = workbook.createSheet("Table");

        // Create a Font for styling header cells
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 14);
        headerFont.setColor(IndexedColors.RED.getIndex());

        // Create a CellStyle with the font
        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);

        // Create a Row
        Row headerRow = sheet.createRow(0);

        // Create cells
        for(int i = 0; i < columns.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(columns[i]);
            cell.setCellStyle(headerCellStyle);
        }

        // Create Cell Style for formatting Date
        CellStyle dateCellStyle = workbook.createCellStyle();
        dateCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd-MM-yyyy"));

        // Create Other rows and cells with employees data
        int rowNum = 1;
        for(String[] d: datalist) {
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(d[0]);
            row.createCell(1).setCellValue(d[1]);
        }

		// Resize all columns to fit the content size
        for(int i = 0; i < columns.length; i++) {
            sheet.autoSizeColumn(i);
        }

        // Write the output to a file
        FileOutputStream fileOut=null;
        try {
        	fileOut = new FileOutputStream("D:\\Freelancer\\thoan_excel\\Results\\"+fileName+".xlsx");
			workbook.write(fileOut);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}finally {
			if(fileOut !=null) {
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


	public static void main(String s[]) {
		JFrame frame = new JFrame("");
		HtmlParser4 panel = new HtmlParser4();
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