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
import java.nio.file.Files;
import java.awt.*;
import java.util.List;
import java.util.*;

public class ExcelProcessing7 extends JPanel implements ActionListener {
	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;

	JButton go;
	JFileChooser chooser;
	String choosertitle;
	File folder;
	private static Map<String, String> datalist = new HashMap<>();

	public ExcelProcessing7() {
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

				FacebookUser tempData = null;
				for (int i = 1; i < table1Sheet.getPhysicalNumberOfRows(); i++) {
					try {
						String imageLink = table1Sheet.getRow(i).getCell(11).getStringCellValue();
						String sku = table1Sheet.getRow(i).getCell(12).getStringCellValue();

						datalist.put(sku, imageLink);
					} catch (Exception e) {
						// TODO: handle exception
						e.printStackTrace();
					}
				}

				System.out.println("YYYYY" + datalist.size());

			}

			File imageFolder = new File("C:\\Users\\VIET\\Desktop\\Upwork Data Entry PDFs\\Images");
			File[] listOfFiles = imageFolder.listFiles();
			String keyValid = null;
			for (File f : listOfFiles) {
				if(f.getName().equals("a0732fafa5ad6c34a4ce096e03955178.jpg")) {
					System.out.println("test");
				}
				for (Map.Entry<String, String> entry : datalist.entrySet()) {

					try {
						if(entry.getValue().equals("a0732fafa5ad6c34a4ce096e03955178.jpg")) {
							System.out.println("XXXXXXXXXXxx");
						}
						if(entry.getKey().equals("OEM-NIS-295.jpg")) {
							
						}
						if (f.getName().equals(entry.getValue())) {
							String newFileName = "C:\\Users\\VIET\\Desktop\\Upwork Data Entry PDFs\\Images2\\"+ entry.getKey().trim();
							Files.copy(f.toPath(), new File(newFileName).toPath());
							
						}
					} catch (Exception e) {

					}
				}
			}

			for (Map.Entry<String, String> entry : datalist.entrySet()) {
				// System.out.println(entry.getKey()+ "/"+entry.getValue());
			}

			inputFile.close();
			System.out.println("$$$$$$$$$$FINISH$$$$$$$$$$$$$");

		} catch (Exception ioe) {
			ioe.printStackTrace();
		}
	}

	private static void writeTxtFile() {
		try {
			File fout = new File("C:\\Users\\VIET\\Desktop\\Upwork Data Entry PDFs\\imageList.txt");
			FileOutputStream fos = new FileOutputStream(fout);

			BufferedWriter bw = new BufferedWriter(new OutputStreamWriter(fos));

			for (String d : datalist.keySet()) {

				bw.write(d);
				bw.newLine();
			}

			bw.close();
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
	}

	public static void main(String s[]) {
		JFrame frame = new JFrame("");
		ExcelProcessing7 panel = new ExcelProcessing7();
		frame.addWindowListener(new WindowAdapter() {
			public void windowClosing(WindowEvent e) {
				System.exit(0);
			}
		});
		frame.getContentPane().add(panel, "Center");
		frame.setSize(panel.getPreferredSize());
		frame.setVisible(true);
//		 File imageFolder = new File("C:\\Users\\VIET\\Desktop\\Upwork Data Entry PDFs\\Images2");
//		 File[] listOfFiles = imageFolder.listFiles();
//		 for(File f: listOfFiles) {
//		 datalist.put(f.getName(), null);
//		 }
//		 writeTxtFile();

	}
}