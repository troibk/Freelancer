package thoan_excel;

import javax.swing.*;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.awt.event.*;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.awt.*;
import java.util.List;
import java.util.*;

public class ExcelProcessing5 extends JPanel implements ActionListener {
	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;

	JButton go;
	JFileChooser chooser;
	String choosertitle;
	File folder;
    private static List<FirmData2> datalist = new ArrayList<>();

	public ExcelProcessing5() {
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
		List<File> excelFiles = new ArrayList<>();
		if (folder != null && folder.isDirectory()) {
			File[] listOfFiles = folder.listFiles();
			String fileExt = "";
			for (File f : listOfFiles) {
				if (!f.isHidden()) {
					fileExt = f.getName();
					if (fileExt.endsWith(".xls") || fileExt.endsWith(".XLS") || fileExt.endsWith(".xlsm")
							|| fileExt.endsWith(".xlsx") || fileExt.endsWith(".XLSM") || fileExt.endsWith(".XLSX")) {
						excelFiles.add(f);
					}
				}
			}
		}
		int i = 0;
		for (File f : excelFiles) {
			processExcel(f);
			i++;
			System.out.println(i);
		}
	}

	public Dimension getPreferredSize() {
		return new Dimension(100, 100);
	}

	public void processExcel(File file) {
		try {
			FileInputStream inputFile = new FileInputStream(file);
			XSSFWorkbook wb = new XSSFWorkbook(inputFile);
			XSSFSheet table1Sheet = wb.getSheetAt(0);

			if (table1Sheet == null) {
				System.out.println("KO CO SHEET :" + table1Sheet);
			} else {
				System.out.println("NO Rows = "+ table1Sheet.getPhysicalNumberOfRows());
				FirmData2 data=null;
				for (int i = 0; i < 3110; i++) {
					try {
						Row row= table1Sheet.getRow(i);
						if(row !=null) {
						String name = table1Sheet.getRow(i).getCell(0).getStringCellValue();
						if(name==null | name.isEmpty()) {
							continue;
						}
						Cell emailCell=table1Sheet.getRow(i).getCell(1);
						String email=null;
						if(emailCell!=null) {
							email = emailCell.getStringCellValue();
						}
						data=new FirmData2(name,email);
						datalist.add(data);
						}
					} catch (Exception e) {
						System.out.println("XXXXXXX" + i);
						if(i==58) {
							e.printStackTrace();
						}
					}

				}

				System.out.println("YYYYY" + datalist.size());
				writeExcelFile();
				
			}

			inputFile.close();
			System.out.println("$$$$$$$$$$FINISH$$$$$$$$$$$$$");

		} catch (Exception ioe) {
			ioe.printStackTrace();
		}
	}
	
	
	private void writeExcelFile() {
		String[] columns = {"Company", "Email"};
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
        for(FirmData2 d: datalist) {
            Row row = sheet.createRow(rowNum++);

            row.createCell(0)
                    .setCellValue(d.getName());

            row.createCell(1)
                    .setCellValue(d.getEmail());
        }

		// Resize all columns to fit the content size
        for(int i = 0; i < columns.length; i++) {
            sheet.autoSizeColumn(i);
        }

        // Write the output to a file
        FileOutputStream fileOut=null;
        try {
        	fileOut = new FileOutputStream("D:\\Freelancer\\thoan_excel\\Results\\20190531.xlsx");
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
		ExcelProcessing5 panel = new ExcelProcessing5();
		frame.addWindowListener(new WindowAdapter() {
			public void windowClosing(WindowEvent e) {
				System.exit(0);
			}
		});
		frame.getContentPane().add(panel, "Center");
		frame.setSize(panel.getPreferredSize());
		frame.setVisible(true);
	}
}