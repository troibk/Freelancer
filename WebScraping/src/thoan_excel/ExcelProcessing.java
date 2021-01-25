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

public class ExcelProcessing extends JPanel implements ActionListener {
	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;

	JButton go;
	JFileChooser chooser;
	String choosertitle;
	File folder;
	private static List<String[]> datalist = new ArrayList<>();

	public ExcelProcessing() {
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
				for (int i = 2; i < table1Sheet.getPhysicalNumberOfRows(); i++) {
					try {
						String item_make = table1Sheet.getRow(i).getCell(3).getStringCellValue();
						String item_model = table1Sheet.getRow(i).getCell(4).getStringCellValue();
						int start_year = (int) table1Sheet.getRow(i).getCell(1).getNumericCellValue();
						if(table1Sheet.getRow(i).getCell(2)==null || table1Sheet.getRow(i).getCell(2).getCellType()==Cell.CELL_TYPE_STRING) {
							datalist.add(new String[] {""+start_year, item_make, item_model});
						}else {
							int end_year = (int) table1Sheet.getRow(i).getCell(2).getNumericCellValue();
							for(int v= start_year; v<=end_year;v++) {
								datalist.add(new String[] {""+v, item_make, item_model});
							}
						}
						if(i%100==0) {
							appendExcelFile("sample_product_ymm.xlsx", "");
							System.out.println("Write==="+i);
						}

					}catch (Exception e) {
						e.printStackTrace();
						System.out.println("EEE "+i);
					}
				
			}
			
			appendExcelFile("sample_product_ymm.xlsx", "");
			inputFile.close();
			System.out.println("$$$$$$$$$$FINISH$$$$$$$$$$$$$");
			}

		} catch (Exception ioe) {
			ioe.printStackTrace();
			
		}
	}

	private void appendExcelFile(String fileName, String sheetName) {
		Workbook workbook = null;
		Sheet sheet;
		try {
			File file = new File("D:\\Freelancer\\thoan_excel\\Results\\" + fileName);
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

		// Create Other rows and cells with employees data

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
		ExcelProcessing panel = new ExcelProcessing();
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