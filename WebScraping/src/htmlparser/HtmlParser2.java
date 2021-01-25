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
 


public class HtmlParser2 extends JPanel implements ActionListener{
	private static final long serialVersionUID = 1L;

	JButton go;
	JFileChooser chooser;
	String choosertitle;
	File folder;
    private static List<Model> datalist = new ArrayList<>();

	public HtmlParser2() {
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
		writeExcelFile("ILCO");
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
			 Element makeE= doc.selectFirst("div.well");
			 String make_key = makeE.text();
			 String[] temp = make_key.split("-");
			 String make="";
			 String keyType="";
			 if(temp.length==3) {
			  make = temp[1].split("Type")[0];
			  keyType= temp[2];
			 }else if(temp.length==2) {
				 keyType= temp[1];
			 }
			 Elements product_box = doc.select("div.product-box");
			 
			 for (Element element: product_box) {
				Model model = new Model();
				model.setKeyType(keyType);
				model.setMake(make);
				String modelName = element.select("a.product-title").text(); 
				if(make!=null && !make.isEmpty()) {
				String[] temps=modelName.split(make);
				
				if(temps.length>1) {
					model.setModel(temps[1]);
				}else {
					model.setModel(modelName);
					System.out.println("+++++E: "+ modelName);
				}
				String[] times= temps[0].split("-");
				String dateFrom = times[0];
				model.setDateFrom(dateFrom);
				if(times.length>1) {
					String dateTo = times[1];
					model.setDateTo(dateTo);
				}
				}else {
					model.setModel(modelName);
				}
				
				Elements p= element.getElementsByTag("p");
				String note="";
				for(Element e : p) {
					String value = e.text();
					if(value.toLowerCase().trim().startsWith("price:")) {
						model.setPrice(value);
					}else if(value.toLowerCase().trim().startsWith("oe:")) {
						model.setOe(value);
					}else if(value.toLowerCase().trim().startsWith("sku:")) {
						model.setSku(value);
					}else if(value.toLowerCase().trim().startsWith("fcc:")) {
						model.setFcc(value);
					}else if(value.toLowerCase().trim().startsWith("chip id:")) {
						
					}else if(value.toLowerCase().trim().startsWith("battery:")) {
						
					}else {
						note=note+value+"$";
					}
				}
				
			
				model.setNote(note);
				
				 Element imageElement = element.select("img").first();

				 String srcValue = imageElement.attr("src");
				 
				 model.setImage(srcValue);
				
				datalist.add(model);
			 }

			

		} catch (Exception ioe) {
			ioe.printStackTrace();
		}

		
	}
	
	
	private void writeExcelFile(String fileName) {
		String[] columns = {"Year From", "Year To","Key Type", "Make","Model", "Price", "SKU", "FCC", "OE","Note","Image"};
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
        for(Model d: datalist) {
            Row row = sheet.createRow(rowNum++);
            
            row.createCell(0).setCellValue(d.getDateFrom());
            row.createCell(1).setCellValue(d.getDateTo());
            
            row.createCell(2).setCellValue(d.getKeyType());
            row.createCell(3).setCellValue(d.getMake());
            
            row.createCell(4).setCellValue(d.getModel());
            row.createCell(5).setCellValue(d.getPrice());
            row.createCell(6).setCellValue(d.getSku());

            row.createCell(7).setCellValue(d.getFcc());
            row.createCell(8).setCellValue(d.getOe());
            row.createCell(9).setCellValue(d.getNote());
            row.createCell(10).setCellValue(d.getImage());
            
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
		HtmlParser2 panel = new HtmlParser2();
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