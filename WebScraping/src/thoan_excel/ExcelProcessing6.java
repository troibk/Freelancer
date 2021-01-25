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

public class ExcelProcessing6 extends JPanel implements ActionListener {
	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;

	JButton go;
	JFileChooser chooser;
	String choosertitle;
	File folder;
	private static List<Map<String, Object>> dataList = new ArrayList<>();

	public ExcelProcessing6() {
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
			dataList.clear();
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
			String fileName = file.getName();
			for (int s = 0; s < wb.getNumberOfSheets(); s++) {
				XSSFSheet table1Sheet = wb.getSheetAt(s);
				String sheetname = table1Sheet.getSheetName();
				
				if (table1Sheet == null) {
					System.out.println("KO CO SHEET :" + table1Sheet);
				} else {
					System.out.println("NO Rows = " + table1Sheet.getPhysicalNumberOfRows());
					
					Map<String, Object> data = null;
					boolean isDateRow = false;
					boolean isModelRow = false;
					for (int j = 0; j < 3; j++) {
						for (int i = 0; i < table1Sheet.getPhysicalNumberOfRows(); i++) {
							try {
								Row row = table1Sheet.getRow(i);
								if (row != null) {
									Cell cell = row.getCell(j);
									if (cell != null) {
										String rowData = "";
										if (cell.getCellType() == Cell.CELL_TYPE_BLANK) {
											continue;
										} else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
											rowData = rowData + cell.getNumericCellValue();
											if(rowData.contains(".")) {
												rowData=rowData.split("\\.")[0];
											}
										} else {
											rowData = cell.getStringCellValue();
										}
										rowData=rowData.trim();
										if (!rowData.isEmpty()) {

											if (rowData.contains("$$$$$$$")) {
												if (data != null) {
													dataList.add(data);
												}
												
												data = createMap();
												data.put(Constant.Page_Number, sheetname);
												isDateRow = true;
												continue;

											} else {
												if (isDateRow) {
													String[] dates = rowData.split("-");
													data.put(Constant.Year_From, dates[0]);
													if (dates.length > 1) {
														if (dates[1].length() < 3) {
															data.put(Constant.Year_To, "20" + dates[1]);
														} else {
															data.put(Constant.Year_To, dates[1]);

														}
													}
													data.put(Constant.Make, "Infinity");
													isDateRow = false;
													isModelRow = true;
													continue;
												} else if (isModelRow) {
													int lastIndex = rowData.lastIndexOf("(");
													if(lastIndex>0) {
													data.put(Constant.Model, rowData.substring(0, lastIndex - 1));
													
														data.put(Constant.System_Type, rowData.substring(lastIndex + 1,
																rowData.trim().length() - 1));
													
													}else {
														data.put(Constant.Model, rowData);
													}
													
													isModelRow = false;
													continue;
													// }else if(rowData.startsWith("A-TEK")) {
													// data.put(Constant.A_TEK, rowData.split(":")[1].trim());
													// }else if(rowData.startsWith("BARNES")) {
													// List<String> listData=(List)data.get(Constant.BARNES);
													// listData.add(rowData.split(":")[1].trim());
													// }else if(rowData.startsWith("HATA")) {
													// List<String> listData=(List)data.get(Constant.HATA);
													// listData.add(rowData.split(":")[1].trim());
													// }else if(rowData.startsWith("HILLMAN")) {
													// data.put(Constant.HILLMAN, rowData.split(":")[1].trim());
													// }else if(rowData.startsWith("HOWARD KEYS")) {
													// data.put(Constant.HOWARD_KEYS, rowData.split(":")[1].trim());
													// }else if(rowData.startsWith("HYKO")) {
													// data.put(Constant.HOWARD_KEYS, rowData.split(":")[1].trim());
												} else {
													List<String> listData = new ArrayList<>();
													
													if (rowData.startsWith(Constant.EMERGENCY_KEY) | rowData.startsWith("EMERG. KEY")) {
														listData = (List) data.get(Constant.EMERGENCY_KEY);
														String[] tempList = {};
														if (rowData.contains(":")) {
															tempList = rowData.split(":")[1].trim().split(",");
														} else if (rowData.contains(";")) {
															tempList = rowData.split(";")[1].trim().split(",");
														} else if (rowData.contains("#")) {
															tempList = rowData.split("#")[1].trim().split(",");
														} else if (rowData.contains("=")) {
															tempList = rowData.split("=")[1].trim().split(",");
														}else {
															System.out.println("EEEEEEEEEEEEEE"+rowData+","+sheetname+","+"C"+j+",R"+i);
														}

														for (String d : tempList) {
															listData.add(d.trim());
														}
														
													
													}else if(rowData.startsWith("ILCO E-KEY:")) {
														listData = (List) data.get(Constant.ILCO);
														String[] tempList = {};
														if (rowData.contains(":")) {
															tempList = rowData.split(":")[1].trim().split(",");
														} else if (rowData.contains(";")) {
															tempList = rowData.split(";")[1].trim().split(",");
														} else if (rowData.contains("#")) {
															tempList = rowData.split("#")[1].trim().split(",");
														} else if (rowData.contains("=")) {
															tempList = rowData.split("=")[1].trim().split(",");
														}else {
															System.out.println("EEEEEEEEEEEEEE"+rowData+","+sheetname+","+"C"+j+",R"+i);
														}

														for (String d : tempList) {
															listData.add(d.trim()+ (" (E-KEY)"));
														}
													}else if(rowData.startsWith("STRATTEC E-KEY")) {
														listData = (List) data.get(Constant.STRATTEC);
														String[] tempList = {};
														if (rowData.contains(":")) {
															tempList = rowData.split(":")[1].trim().split(",");
														} else if (rowData.contains(";")) {
															tempList = rowData.split(";")[1].trim().split(",");
														} else if (rowData.contains("#")) {
															tempList = rowData.split("#")[1].trim().split(",");
														} else if (rowData.contains("=")) {
															tempList = rowData.split("=")[1].trim().split(",");
														}else {
															System.out.println("EEEEEEEEEEEEEE"+rowData+","+sheetname+","+"C"+j+",R"+i);
														}

														for (String d : tempList) {
															listData.add(d.trim() +" (E-KEY)");
														}
													}else if(rowData.startsWith("STRATTEC E-KEY")) {
														listData = (List) data.get(Constant.STRATTEC);
														String[] tempList = {};
														if (rowData.contains(":")) {
															tempList = rowData.split(":")[1].trim().split(",");
														} else if (rowData.contains(";")) {
															tempList = rowData.split(";")[1].trim().split(",");
														} else if (rowData.contains("#")) {
															tempList = rowData.split("#")[1].trim().split(",");
														} else if (rowData.contains("=")) {
															tempList = rowData.split("=")[1].trim().split(",");
														}else {
															System.out.println("EEEEEEEEEEEEEE"+rowData+","+sheetname+","+"C"+j+",R"+i);
														}

														for (String d : tempList) {
															listData.add(d.trim() +" (E-KEY)");
														}
													}else if(rowData.startsWith(Constant.ILCO)) {
														listData = (List) data.get(Constant.ILCO);
														String[] tempList = {};
														if (rowData.contains(":")) {
															tempList = rowData.split(":")[1].trim().split(",");
														} else if (rowData.contains(";")) {
															tempList = rowData.split(";")[1].trim().split(",");
														} else if (rowData.contains("#")) {
															tempList = rowData.split("#")[1].trim().split(",");
														} else if (rowData.contains("=")) {
															tempList = rowData.split("=")[1].trim().split(",");
														}else {
															System.out.println("EEEEEEEEEEEEEE"+rowData+","+sheetname+","+"C"+j+",R"+i);
														}

														for (String d : tempList) {
															listData.add(d.trim());
														}
													} else if (rowData.startsWith(Constant.JET)) {
														listData = (List) data.get(Constant.JET);
														String[] tempList = {};
														if (rowData.contains(":")) {
															tempList = rowData.split(":")[1].trim().split(",");
														} else if (rowData.contains(";")) {
															tempList = rowData.split(";")[1].trim().split(",");
														} else if (rowData.contains("#")) {
															tempList = rowData.split("#")[1].trim().split(",");
														} else if (rowData.contains("=")) {
															tempList = rowData.split("=")[1].trim().split(",");
														}else {
															System.out.println("EEEEEEEEEEEEEE"+rowData+","+sheetname+","+"C"+j+",R"+i);
														}

														for (String d : tempList) {
															listData.add(d.trim());
														}
													} else if (rowData.startsWith(Constant.JMA)) {
														listData = (List) data.get(Constant.JMA);
														String[] tempList = {};
														if (rowData.contains(":")) {
															tempList = rowData.split(":")[1].trim().split(",");
														} else if (rowData.contains(";")) {
															tempList = rowData.split(";")[1].trim().split(",");
														} else if (rowData.contains("#")) {
															tempList = rowData.split("#")[1].trim().split(",");
														} else if (rowData.contains("=")) {
															tempList = rowData.split("=")[1].trim().split(",");
														}else {
															System.out.println("EEEEEEEEEEEEEE"+rowData+","+sheetname+","+"C"+j+",R"+i);
														}

														for (String d : tempList) {
															listData.add(d.trim());
														}
													} else if (rowData.toUpperCase().startsWith(Constant.KEYLINE)) {
														listData = (List) data.get(Constant.KEYLINE);
														String[] tempList = {};
														if (rowData.contains(":")) {
															tempList = rowData.split(":")[1].trim().split(",");
														} else if (rowData.contains(";")) {
															tempList = rowData.split(";")[1].trim().split(",");
														} else if (rowData.contains("#")) {
															tempList = rowData.split("#")[1].trim().split(",");
														} else if (rowData.contains("=")) {
															tempList = rowData.split("=")[1].trim().split(",");
														}else {
															System.out.println("EEEEEEEEEEEEEE"+rowData+","+sheetname+","+"C"+j+",R"+i);
														}

														for (String d : tempList) {
															listData.add(d.trim());
														}
													} else if (rowData.toUpperCase().startsWith(Constant.OEM_Remote)) {
														listData = (List) data.get(Constant.OEM);
														String[] tempList = {};
														if (rowData.contains(":")) {
															tempList = rowData.split(":")[1].trim().split(",");
														} else if (rowData.contains(";")) {
															tempList = rowData.split(";")[1].trim().split(",");
														} else if (rowData.contains("#")) {
															tempList = rowData.split("#")[1].trim().split(",");
														} else if (rowData.contains("=")) {
															tempList = rowData.split("=")[1].trim().split(",");
														}else {
															System.out.println("EEEEEEEEEEEEEE"+rowData+","+sheetname+","+"C"+j+",R"+i);
														}

														for (String d : tempList) {
															listData.add(d.trim()+" (Remote)");
														}
													} else if (rowData.toUpperCase().startsWith(Constant.OEM_COMPATIBLE)) {
														listData = (List) data.get(Constant.OEM);
														String[] tempList = {};
														if (rowData.contains(":")) {
															tempList = rowData.split(":")[1].trim().split(",");
														} else if (rowData.contains(";")) {
															tempList = rowData.split(";")[1].trim().split(",");
														} else if (rowData.contains("#")) {
															tempList = rowData.split("#")[1].trim().split(",");
														} else if (rowData.contains("=")) {
															tempList = rowData.split("=")[1].trim().split(",");
														}else {
															System.out.println("EEEEEEEEEEEEEE"+rowData+","+sheetname+","+"C"+j+",R"+i);
														}

														for (String d : tempList) {
															listData.add(d.trim()+" (Compatible)");
														}

													} else if (rowData.toUpperCase().startsWith(Constant.OEM_Emergency)) {
														listData = (List) data.get(Constant.OEM);
														String[] tempList = {};
														if (rowData.contains(":")) {
															tempList = rowData.split(":")[1].trim().split(",");
														} else if (rowData.contains(";")) {
															tempList = rowData.split(";")[1].trim().split(",");
														} else if (rowData.contains("#")) {
															tempList = rowData.split("#")[1].trim().split(",");
														} else if (rowData.contains("=")) {
															tempList = rowData.split("=")[1].trim().split(",");
														}else {
															System.out.println("EEEEEEEEEEEEEE"+rowData+","+sheetname+","+"C"+j+",R"+i);
														}

														for (String d : tempList) {
															listData.add(d.trim()+" (EmergencyKey)");
														}
													} else if (rowData.toUpperCase().startsWith(Constant.OEM_Key_blade)) {
														listData = (List) data.get(Constant.OEM);
														String[] tempList = {};
														if (rowData.contains(":")) {
															tempList = rowData.split(":")[1].trim().split(",");
														} else if (rowData.contains(";")) {
															tempList = rowData.split(";")[1].trim().split(",");
														} else if (rowData.contains("#")) {
															tempList = rowData.split("#")[1].trim().split(",");
														} else if (rowData.contains("=")) {
															tempList = rowData.split("=")[1].trim().split(",");
														}else {
															System.out.println("EEEEEEEEEEEEEE"+rowData+","+sheetname+","+"C"+j+",R"+i);
														}

														for (String d : tempList) {
															listData.add(d.trim()+" (Key blade)");
														}
													} else if (rowData.toUpperCase().startsWith(Constant.OEM_Roll_pin)) {
														listData = (List) data.get(Constant.OEM);
														String[] tempList = {};
														if (rowData.contains(":")) {
															tempList = rowData.split(":")[1].trim().split(",");
														} else if (rowData.contains(";")) {
															tempList = rowData.split(";")[1].trim().split(",");
														} else if (rowData.contains("#")) {
															tempList = rowData.split("#")[1].trim().split(",");
														} else if (rowData.contains("=")) {
															tempList = rowData.split("=")[1].trim().split(",");
														}else {
															System.out.println("EEEEEEEEEEEEEE"+rowData+","+sheetname+","+"C"+j+",R"+i);
														}

														for (String d : tempList) {
															listData.add(d.trim()+" (Roll pin)");
														}

													}else if (rowData.toUpperCase().startsWith(Constant.OEM_Ekey) | rowData.toUpperCase().startsWith("E-KEY:")) {
														listData = (List) data.get(Constant.OEM);
														String[] tempList = {};
														if (rowData.contains(":")) {
															tempList = rowData.split(":")[1].trim().split(",");
														} else if (rowData.contains(";")) {
															tempList = rowData.split(";")[1].trim().split(",");
														} else if (rowData.contains("#")) {
															tempList = rowData.split("#")[1].trim().split(",");
														} else if (rowData.contains("=")) {
															tempList = rowData.split("=")[1].trim().split(",");
														}else {
															System.out.println("EEEEEEEEEEEEEE"+rowData+","+sheetname+","+"C"+j+",R"+i);
														}

														for (String d : tempList) {
															listData.add(d.trim()+" (E-key)");
														}
													} else if (rowData.toUpperCase().startsWith(Constant.OEM_PROX) | rowData.toUpperCase().startsWith("PROX:")) {
														listData = (List) data.get(Constant.OEM);
														String[] tempList = {};
														if (rowData.contains(":")) {
															tempList = rowData.split(":")[1].trim().split(",");
														} else if (rowData.contains(";")) {
															tempList = rowData.split(";")[1].trim().split(",");
														} else if (rowData.contains("#")) {
															tempList = rowData.split("#")[1].trim().split(",");
														} else if (rowData.contains("=")) {
															tempList = rowData.split("=")[1].trim().split(",");
														}else {
															System.out.println("EEEEEEEEEEEEEE"+rowData+","+sheetname+","+"C"+j+",R"+i);
														}

														for (String d : tempList) {
															listData.add(d.trim()+" (Prox)");
														}	
													} else if (rowData.toUpperCase().startsWith(Constant.OEM_FLIP)) {
														listData = (List) data.get(Constant.OEM);
														String[] tempList = {};
														if (rowData.contains(":")) {
															tempList = rowData.split(":")[1].trim().split(",");
														} else if (rowData.contains(";")) {
															tempList = rowData.split(";")[1].trim().split(",");
														} else if (rowData.contains("#")) {
															tempList = rowData.split("#")[1].trim().split(",");
														} else if (rowData.contains("=")) {
															tempList = rowData.split("=")[1].trim().split(",");
														}else {
															System.out.println("EEEEEEEEEEEEEE"+rowData+","+sheetname+","+"C"+j+",R"+i);
														}

														for (String d : tempList) {
															listData.add(d.trim()+" (Flip Key)");
														}	
													} else if (rowData.toUpperCase().startsWith(Constant.OEM_Smart)) {
														listData = (List) data.get(Constant.OEM);
														String[] tempList = {};
														if (rowData.contains(":")) {
															tempList = rowData.split(":")[1].trim().split(",");
														} else if (rowData.contains(";")) {
															tempList = rowData.split(";")[1].trim().split(",");
														} else if (rowData.contains("#")) {
															tempList = rowData.split("#")[1].trim().split(",");
														} else if (rowData.contains("=")) {
															tempList = rowData.split("=")[1].trim().split(",");
														}else {
															System.out.println("EEEEEEEEEEEEEE"+rowData+","+sheetname+","+"C"+j+",R"+i);
														}

														for (String d : tempList) {
															listData.add(d.trim()+" (SmartKey)");
														}	
													}else if (rowData.toUpperCase().startsWith(Constant.OEM)) {
														listData = (List) data.get(Constant.OEM);
														String[] tempList = {};
														if (rowData.contains(":")) {
															tempList = rowData.split(":")[1].trim().split(",");
														} else if (rowData.contains(";")) {
															tempList = rowData.split(";")[1].trim().split(",");
														} else if (rowData.contains("#")) {
															tempList = rowData.split("#")[1].trim().split(",");
														} else if (rowData.contains("=")) {
															tempList = rowData.split("=")[1].trim().split(",");
														}else {
															System.out.println("EEEEEEEEEEEEEE"+rowData+","+sheetname+","+"C"+j+",R"+i);
														}

														for (String d : tempList) {
															listData.add(d.trim());
														}
													}else if (rowData.toUpperCase().startsWith(Constant.PCB)) {
														listData = (List) data.get(Constant.PCB);
														String[] tempList = {};
														if (rowData.contains(":")) {
															tempList = rowData.split(":")[1].trim().split(",");
														} else if (rowData.contains(";")) {
															tempList = rowData.split(";")[1].trim().split(",");
														} else if (rowData.contains("#")) {
															tempList = rowData.split("#")[1].trim().split(",");
														} else if (rowData.contains("=")) {
															tempList = rowData.split("=")[1].trim().split(",");
														}else {
															System.out.println("EEEEEEEEEEEEEE"+rowData+","+sheetname+","+"C"+j+",R"+i);
														}

														for (String d : tempList) {
															listData.add(d.trim());
														}
													} else if (rowData.toUpperCase().startsWith(Constant.IC)) {
														listData = (List) data.get(Constant.IC);
														String[] tempList = {};
														if (rowData.contains(":")) {
															tempList = rowData.split(":")[1].trim().split(",");
														} else if (rowData.contains(";")) {
															tempList = rowData.split(";")[1].trim().split(",");
														} else if (rowData.contains("#")) {
															tempList = rowData.split("#")[1].trim().split(",");
														} else if (rowData.contains("=")) {
															tempList = rowData.split("=")[1].trim().split(",");
														}else {
															System.out.println("EEEEEEEEEEEEEE"+rowData+","+sheetname+","+"C"+j+",R"+i);
														}

														for (String d : tempList) {
															listData.add(d.trim());
														}
													} else if (rowData.toUpperCase().startsWith(Constant.STRATTEC_REMOTE)
															| rowData.toUpperCase().startsWith("STRATIEC REMOTE")) {
														listData = (List) data.get(Constant.STRATTEC);
														String[] tempList = {};
														if (rowData.contains(":")) {
															tempList = rowData.split(":")[1].trim().split(",");
														} else if (rowData.contains(";")) {
															tempList = rowData.split(";")[1].trim().split(",");
														} else if (rowData.contains("#")) {
															tempList = rowData.split("#")[1].trim().split(",");
														} else if (rowData.contains("=")) {
															tempList = rowData.split("=")[1].trim().split(",");
														}else {
															System.out.println("EEEEEEEEEEEEEE"+rowData+","+sheetname+","+"C"+j+",R"+i);
														}

														for (String d : tempList) {
															listData.add(d.trim()+ " (Remote)");
														}
													} else if (rowData.toUpperCase().startsWith(Constant.STRATTEC_COMPATIBLE)
															| rowData.toUpperCase().startsWith("COMPATIBLE STRATIEC")) {
														listData = (List) data.get(Constant.STRATTEC);
														String[] tempList = {};
														if (rowData.contains(":")) {
															tempList = rowData.split(":")[1].trim().split(",");
														} else if (rowData.contains(";")) {
															tempList = rowData.split(";")[1].trim().split(",");
														} else if (rowData.contains("#")) {
															tempList = rowData.split("#")[1].trim().split(",");
														} else if (rowData.contains("=")) {
															tempList = rowData.split("=")[1].trim().split(",");
														}else {
															System.out.println("EEEEEEEEEEEEEE"+rowData+","+sheetname+","+"C"+j+",R"+i);
														}

														for (String d : tempList) {
															listData.add(d.trim()+ " (Compatible)");
														}
													} else if (rowData.toUpperCase().startsWith(Constant.STRATTEC)
															| rowData.toUpperCase().startsWith("STRATIEC")) {
														listData = (List) data.get(Constant.STRATTEC);
														String[] tempList = {};
														if (rowData.contains(":")) {
															tempList = rowData.split(":")[1].trim().split(",");
														} else if (rowData.contains(";")) {
															tempList = rowData.split(";")[1].trim().split(",");
														} else if (rowData.contains("#")) {
															tempList = rowData.split("#")[1].trim().split(",");
														} else if (rowData.contains("=")) {
															tempList = rowData.split("=")[1].trim().split(",");
														}else {
															System.out.println("EEEEEEEEEEEEEE"+rowData+","+sheetname+","+"C"+j+",R"+i);
														}

														for (String d : tempList) {
															listData.add(d.trim());
														}
													}else if (rowData.toUpperCase().startsWith("LISHI:")) {
															listData = (List) data.get(Constant.LISHI);
															String[] tempList = {};
															if (rowData.contains(":")) {
																tempList = rowData.split(":")[1].trim().split(",");
															} else if (rowData.contains(";")) {
																tempList = rowData.split(";")[1].trim().split(",");
															} else if (rowData.contains("#")) {
																tempList = rowData.split("#")[1].trim().split(",");
															} else if (rowData.contains("=")) {
																tempList = rowData.split("=")[1].trim().split(",");
															}else {
																System.out.println("EEEEEEEEEEEEEE"+rowData+","+sheetname+","+"C"+j+",R"+i);
															}

															for (String d : tempList) {
																listData.add(d.trim());
															}
													}
													else if (rowData.toUpperCase().startsWith(Constant.LISHI21)
															| rowData.startsWith("LISH121") | rowData.startsWith("L1SHI21") | rowData.startsWith("LISHI21") | rowData.startsWith("LISI2")) {
														listData = (List) data.get(Constant.LISHI21);
														String[] tempList = {};
														if (rowData.contains(":")) {
															tempList = rowData.split(":")[1].trim().split(",");
														} else if (rowData.contains(";")) {
															tempList = rowData.split(";")[1].trim().split(",");
														} else if (rowData.contains("#")) {
															tempList = rowData.split("#")[1].trim().split(",");
														} else if (rowData.contains("=")) {
															tempList = rowData.split("=")[1].trim().split(",");
														}else {
															System.out.println("EEEEEEEEEEEEEE"+rowData+","+sheetname+","+"C"+j+",R"+i);
														}

														for (String d : tempList) {
															listData.add(d.trim());
														}
													} else if (rowData.toUpperCase().startsWith("MECHANICAL")) {
														listData = (List) data.get(Constant.MECHANICAL_KEY);
														String[] tempList = {};
														if (rowData.contains(":")) {
															tempList = rowData.split(":")[1].trim().split(",");
														} else if (rowData.contains(";")) {
															tempList = rowData.split(";")[1].trim().split(",");
														} else if (rowData.contains("#")) {
															tempList = rowData.split("#")[1].trim().split(",");
														} else if (rowData.contains("=")) {
															tempList = rowData.split("=")[1].trim().split(",");
														}else {
															System.out.println("EEEEEEEEEEEEEE"+rowData+","+sheetname+","+"C"+j+",R"+i);
														}

														for (String d : tempList) {
															listData.add(d.trim());
														}
													} else if (rowData.startsWith("CODE")) {
														listData = (List) data.get(Constant.CODE_SERIES);
														String[] tempList = {};
														if (rowData.contains(":")) {
															tempList = rowData.split(":")[1].trim().split(",");
														} else if (rowData.contains(";")) {
															tempList = rowData.split(";")[1].trim().split(",");
														} else if (rowData.contains("#")) {
															tempList = rowData.split("#")[1].trim().split(",");
														} else if (rowData.contains("=")) {
															tempList = rowData.split("=")[1].trim().split(",");
														}else {
															System.out.println("EEEEEEEEEEEEEE"+rowData+","+sheetname+","+"C"+j+",R"+i);
														}

														for (String d : tempList) {
															listData.add(d.trim());
														}
													} else if (rowData.startsWith("MAX ")) {
														listData = (List) data.get(Constant.MAX_KEYS);
														String[] tempList = {};
														if (rowData.contains(":")) {
															tempList = rowData.split(":")[1].trim().split(",");
														} else if (rowData.contains(";")) {
															tempList = rowData.split(";")[1].trim().split(",");
														} else if (rowData.contains("#")) {
															tempList = rowData.split("#")[1].trim().split(",");
														} else if (rowData.contains("=")) {
															tempList = rowData.split("=")[1].trim().split(",");
														}else {
															System.out.println("EEEEEEEEEEEEEE"+rowData+","+sheetname+","+"C"+j+",R"+i);
														}

														for (String d : tempList) {
															listData.add(d.trim());
														}
													
													} else if (rowData.toUpperCase().startsWith("MASTER TRANSPONDER")) {
														listData = (List) data.get(Constant.MASTER_TRANSPONDER_DATA);
														String[] tempList = {};
														if (rowData.contains(":")) {
															tempList = rowData.split(":")[1].trim().split(",");
														} else if (rowData.contains(";")) {
															tempList = rowData.split(";")[1].trim().split(",");
														} else if (rowData.contains("#")) {
															tempList = rowData.split("#")[1].trim().split(",");
														} else if (rowData.contains("=")) {
															tempList = rowData.split("=")[1].trim().split(",");
														}else {
															System.out.println("EEEEEEEEEEEEEE"+rowData+","+sheetname+","+"C"+j+",R"+i);
														}

														for (String d : tempList) {
															listData.add(d.trim());
														}
													} else if (rowData.toUpperCase().startsWith("VALET TRANSPONDER")) {
														if (rowData.contains(":")) {
															data.put(Constant.VALET_TRANSPONDER_DATA,rowData.split(":")[1].trim());
														} else if (rowData.contains(";")) {
															data.put(Constant.VALET_TRANSPONDER_DATA,rowData.split(";")[1].trim());
														} else if (rowData.contains("#")) {
															data.put(Constant.VALET_TRANSPONDER_DATA,rowData.split("#")[1].trim());
														} else if (rowData.contains("=")) {
															data.put(Constant.VALET_TRANSPONDER_DATA,rowData.split("=")[1].trim());
														}
														
													} else if (rowData.toUpperCase().startsWith(Constant.PAGE_2_DATA)) {
														if (rowData.contains(":")) {
															data.put(Constant.PAGE_2_DATA,rowData.split(":")[1].trim());
														} else if (rowData.contains(";")) {
															data.put(Constant.PAGE_2_DATA,rowData.split(";")[1].trim());
														} else if (rowData.contains("#")) {
															data.put(Constant.PAGE_2_DATA,rowData.split("#")[1].trim());
														} else if (rowData.contains("=")) {
															data.put(Constant.PAGE_2_DATA,rowData.split("=")[1].trim());
														}
													} else if (rowData.toUpperCase().startsWith("TRANSPONDER ")) {
														if (rowData.contains(":")) {
															data.put(Constant.TRANSPONDER_RE_USE,rowData.split(":")[1].trim());
														} else if (rowData.contains(";")) {
															data.put(Constant.TRANSPONDER_RE_USE,rowData.split(";")[1].trim());
														} else if (rowData.contains("#")) {
															data.put(Constant.TRANSPONDER_RE_USE,rowData.split("#")[1].trim());
														} else if (rowData.contains("=")) {
															data.put(Constant.TRANSPONDER_RE_USE,rowData.split("=")[1].trim());
														}
													} else if (rowData.toUpperCase().startsWith(Constant.TRANSPONDER)) {
														listData = (List) data.get(Constant.TRANSPONDER);
														String[] tempList = {};
														if (rowData.contains(":")) {
															tempList = rowData.split(":")[1].trim().split(",");
														} else if (rowData.contains(";")) {
															tempList = rowData.split(";")[1].trim().split(",");
														} else if (rowData.contains("#")) {
															tempList = rowData.split("#")[1].trim().split(",");
														} else if (rowData.contains("=")) {
															tempList = rowData.split("=")[1].trim().split(",");
														}else {
															System.out.println("EEEEEEEEEEEEEE"+rowData+","+sheetname+","+"C"+j+",R"+i);
														}

														for (String d : tempList) {
															listData.add(d.trim());
														}
														
													}else if (rowData.toUpperCase().startsWith(Constant.Remote)) {
														listData = (List) data.get(Constant.Remote);
														String[] tempList = {};
														if (rowData.contains(":")) {
															tempList = rowData.split(":")[1].trim().split(",");
														} else if (rowData.contains(";")) {
															tempList = rowData.split(";")[1].trim().split(",");
														} else if (rowData.contains("#")) {
															tempList = rowData.split("#")[1].trim().split(",");
														} else if (rowData.contains("=")) {
															tempList = rowData.split("=")[1].trim().split(",");
														}else {
															System.out.println("EEEEEEEEEEEEEE"+rowData+","+sheetname+","+"C"+j+",R"+i);
														}

														for (String d : tempList) {
															listData.add(d.trim());
														}
													} else if (rowData.toUpperCase().startsWith(Constant.FCC_COMPATIBLE)) {
														listData = (List) data.get(Constant.FCC);
														String[] tempList = {};
														if (rowData.contains(":")) {
															tempList = rowData.split(":")[1].trim().split(",");
														} else if (rowData.contains(";")) {
															tempList = rowData.split(";")[1].trim().split(",");
														} else if (rowData.contains("#")) {
															tempList = rowData.split("#")[1].trim().split(",");
														} else if (rowData.contains("=")) {
															tempList = rowData.split("=")[1].trim().split(",");
														}else {
															System.out.println("EEEEEEEEEEEEEE"+rowData+","+sheetname+","+"C"+j+",R"+i);
														}

														for (String d : tempList) {
															listData.add(d.trim()+ " (Compatible)");
														}
													} else if (rowData.toUpperCase().startsWith(Constant.FCC)) {
														listData = (List) data.get(Constant.FCC);
														String[] tempList = {};
														if (rowData.contains(":")) {
															tempList = rowData.split(":")[1].trim().split(",");
														} else if (rowData.contains(";")) {
															tempList = rowData.split(";")[1].trim().split(",");
														} else if (rowData.contains("#")) {
															tempList = rowData.split("#")[1].trim().split(",");
														} else if (rowData.contains("=")) {
															tempList = rowData.split("=")[1].trim().split(",");
														}else {
															System.out.println("EEEEEEEEEEEEEE"+rowData+","+sheetname+","+"C"+j+",R"+i);
														}

														for (String d : tempList) {
															listData.add(d.trim());
														}
													} else if (rowData.toUpperCase().startsWith("MHZ")) {
														listData = (List) data.get(Constant.Frequency);
														String[] tempList = {};
														if (rowData.contains(":")) {
															tempList = rowData.split(":")[1].trim().split(",");
														} else if (rowData.contains(";")) {
															tempList = rowData.split(";")[1].trim().split(",");
														} else if (rowData.contains("#")) {
															tempList = rowData.split("#")[1].trim().split(",");
														} else if (rowData.contains("=")) {
															tempList = rowData.split("=")[1].trim().split(",");
														}else {
															System.out.println("EEEEEEEEEEEEEE"+rowData+","+sheetname+","+"C"+j+",R"+i);
														}

														for (String d : tempList) {
															listData.add(d.trim());
														}
													} else {
														listData = (List) data.get(Constant.newCol);
														listData.add("AAAAAAAASSSS"+sheetname+"_CCC"+j+"_RRR"+i+":"+rowData.trim());
														
													}
													
												}
											}
										}
									}
								}
							} catch (Exception e) {
								System.out.println("SSSS"+sheetname+"CCC" + j + ", RRRR" + i);
								
								e.printStackTrace();

							}

						}
					}
				}
				

			}
			writeExcelFile(fileName);
			inputFile.close();
			System.out.println("$$$$$$$$$$FINISH$$$$$$$$$$$$$");

		} catch (

		Exception ioe) {
			ioe.printStackTrace();
		}
	}

	private void writeExcelFile(String fileName) {
		String[] columns = { "Page#", "Year From", "Year To", "Make", "Model", "System Type", "EMERGENCY KEY1", "EMERGENCY KEY2",
				// "A-TEK",
				// "BARNES1",
				// "BARNES2",
				// "BARNES3",
				// "BARNES4",
				// "BARNES5",
				// "HATA1",
				// "HATA2",
				// "HILLMAN",
				// "HOWARD KEYS",
				// "HYKO",
				"ILCO1", "ILCO2", "ILCO3", "JET1", "JET2", "JET3", "JMA1", "JMA2", "JMA3", "KEYLINE1", "KEYLINE2",
				// "LOCKCRAFT1",
				// "LOCKCRAFT2",
				"OEM1", "OEM2", "OEM3", "OEM4", "OEM5", "OEM6", "OEM7", "OEM8", "OEM9", "OEM10", "OEM11",
				"OEM12", "OEM13", "OEM14", "OEM15", "OEM16", "OEM17", "OEM18",
				"OEM19", "OEM20", "OEM21", "OEM22", "OEM23", "PCB1", "PCB2", "PCB3", "PCB4", "PCB5",
				"IC1","IC2", "STRATTEC1", "STRATTEC2", "STRATTEC3", "STRATTEC4", "STRATTEC5", "STRATTEC6",
				// "WINZER1",
				// "WINZER2",
				// "WINZER3",
				// "ACCU",
				// "ACCU-READER",
				// "BTR",
				// "EEZ READER",
				// "KOBRA",
				"LISHI", "LISHI","LISHI21", "LISHI21", "MECHANICAL KEY1", "MECHANICAL KEY2", "CODE SERIES1", "CODE SERIES2", "MAX KEYS1","MAX KEYS2",
				"TRANSPONDER1", "TRANSPONDER2", "MASTER TRANSPONDER DATA1", "MASTER TRANSPONDER DATA2",
				"VALET TRANSPONDER  DATA", "PAGE 2 DATA", "TRANSPONDER RE-USE", "Remote1", "Remote2", "Remote3", "Remote4", "FCC1", "FCC2", "FCC3", "FCC4", "FCC5","FCC6", "FCC7", "FCC8", "FCC9", "FCC10",
				"Frequency1", "Frequency2" };
		Workbook workbook = new XSSFWorkbook(); // new HSSFWorkbook() for generating `.xls` file

		/*
		 * CreationHelper helps us create instances of various things like DataFormat,
		 * Hyperlink, RichTextString etc, in a format (HSSF, XSSF) independent way
		 */
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
		for (int i = 0; i < columns.length; i++) {
			Cell cell = headerRow.createCell(i);
			cell.setCellValue(columns[i]);
			cell.setCellStyle(headerCellStyle);
		}

		// Create Cell Style for formatting Date
		CellStyle dateCellStyle = workbook.createCellStyle();
		dateCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd-MM-yyyy"));

		// Create Other rows and cells with employees data
		int rowNum = 1;
		for (Map<String, Object> d : dataList) {
			Row row = sheet.createRow(rowNum++);
			int i = 0;
			
			row.createCell(i).setCellValue(d.get(Constant.Page_Number).toString());
			i++;
			row.createCell(i).setCellValue(d.get(Constant.Year_From).toString());
			i++;
			row.createCell(i).setCellValue(d.get(Constant.Year_To).toString());
			i++;
			row.createCell(i).setCellValue(d.get(Constant.Make).toString());
			i++;
			row.createCell(i).setCellValue(d.get(Constant.Model).toString());
			i++;
			row.createCell(i).setCellValue(d.get(Constant.System_Type).toString());


			List<String> listData = (List) d.get(Constant.EMERGENCY_KEY);
			i++;
			row.createCell(i).setCellValue(listData.size() > 0 ? listData.get(0) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 1 ? listData.get(1) : "");
			
			
			// List<String> listData= (List)d.get(Constant.BARNES);
			// row.createCell(i).setCellValue(listData.size()>0?listData.get(0):"");
			// row.createCell(i).setCellValue(listData.size()>1?listData.get(1):"");
			// row.createCell(i).setCellValue(listData.size()>2?listData.get(2):"");
			// row.createCell(i).setCellValue(listData.size()>3?listData.get(3):"");
			// row.createCell(i).setCellValue(listData.size()>4?listData.get(4):"");
			//
			// listData= (List)d.get(Constant.HATA);
			// row.createCell(i).setCellValue(listData.size()>0?listData.get(0):"");
			// row.createCell(i).setCellValue(listData.size()>1?listData.get(1):"");
			//
			// row.createCell(i).setCellValue(d.get(Constant.HILLMAN).toString());
			//
			// row.createCell(i).setCellValue(d.get(Constant.HOWARD_KEYS).toString());
			//
			// row.createCell(i).setCellValue(d.get(Constant.HYKO).toString());

			listData = (List) d.get(Constant.ILCO);
			i++;
			row.createCell(i).setCellValue(listData.size() > 0 ? listData.get(0) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 1 ? listData.get(1) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 2 ? listData.get(2) : "");

			listData = (List) d.get(Constant.JET);
			i++;
			row.createCell(i).setCellValue(listData.size() > 0 ? listData.get(0) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 1 ? listData.get(1) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 2 ? listData.get(2) : "");

			listData = (List) d.get(Constant.JMA);
			i++;
			row.createCell(i).setCellValue(listData.size() > 0 ? listData.get(0) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 1 ? listData.get(1) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 2 ? listData.get(2) : "");

			listData = (List) d.get(Constant.KEYLINE);
			i++;
			row.createCell(i).setCellValue(listData.size() > 0 ? listData.get(0) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 1 ? listData.get(1) : "");

			// listData= (List)d.get(Constant.LOCKCRAFT);
			// row.createCell(i).setCellValue(listData.size()>0?listData.get(0):"");
			// row.createCell(i).setCellValue(listData.size()>1?listData.get(1):"");

			listData = (List) d.get(Constant.OEM);
			i++;
			row.createCell(i).setCellValue(listData.size() > 0 ? listData.get(0) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 1 ? listData.get(1) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 2 ? listData.get(2) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 3 ? listData.get(3) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 4 ? listData.get(4) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 5 ? listData.get(5) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 6 ? listData.get(6) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 7 ? listData.get(7) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 8 ? listData.get(8) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 9 ? listData.get(9) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 10 ? listData.get(10) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 11 ? listData.get(11) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 12 ? listData.get(12) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 13 ? listData.get(13) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 14 ? listData.get(14) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 15 ? listData.get(15) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 16 ? listData.get(16) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 17 ? listData.get(17) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 18 ? listData.get(18) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 19 ? listData.get(19) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 20 ? listData.get(20) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 21 ? listData.get(21) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 22 ? listData.get(22) : "");
			
			listData = (List) d.get(Constant.PCB);
			i++;
			row.createCell(i).setCellValue(listData.size() > 0 ? listData.get(0) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 1 ? listData.get(1) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 2 ? listData.get(2) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 3 ? listData.get(3) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 4 ? listData.get(4) : "");

			listData = (List) d.get(Constant.IC);
			i++;
			row.createCell(i).setCellValue(listData.size() > 0 ? listData.get(0) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 1 ? listData.get(1) : "");
			
			listData = (List) d.get(Constant.STRATTEC);
			i++;
			row.createCell(i).setCellValue(listData.size() > 0 ? listData.get(0) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 1 ? listData.get(1) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 2 ? listData.get(2) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 3 ? listData.get(3) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 4 ? listData.get(4) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 5 ? listData.get(5) : "");

			listData = (List) d.get(Constant.LISHI);
			i++;
			row.createCell(i).setCellValue(listData.size() > 0 ? listData.get(0) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 1 ? listData.get(1) : "");


			listData = (List) d.get(Constant.LISHI21);
			i++;
			row.createCell(i).setCellValue(listData.size() > 0 ? listData.get(0) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 1 ? listData.get(1) : "");

			listData = (List) d.get(Constant.MECHANICAL_KEY);
			i++;
			row.createCell(i).setCellValue(listData.size() > 0 ? listData.get(0) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 1 ? listData.get(1) : "");

			listData = (List) d.get(Constant.CODE_SERIES);
			i++;
			row.createCell(i).setCellValue(listData.size() > 0 ? listData.get(0) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 1 ? listData.get(1) : "");

			listData = (List) d.get(Constant.MAX_KEYS);
			i++;
			row.createCell(i).setCellValue(listData.size() > 0 ? listData.get(0) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 1 ? listData.get(1) : "");

			listData = (List) d.get(Constant.TRANSPONDER);
			i++;
			row.createCell(i).setCellValue(listData.size() > 0 ? listData.get(0) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 1 ? listData.get(1) : "");

			listData = (List) d.get(Constant.MASTER_TRANSPONDER_DATA);
			i++;
			row.createCell(i).setCellValue(listData.size() > 0 ? listData.get(0) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 1 ? listData.get(1) : "");
			i++;
			row.createCell(i).setCellValue(d.get(Constant.VALET_TRANSPONDER_DATA).toString());
			i++;
			row.createCell(i).setCellValue(d.get(Constant.PAGE_2_DATA).toString());
			i++;
			row.createCell(i).setCellValue(d.get(Constant.TRANSPONDER_RE_USE).toString());
			
			listData = (List) d.get(Constant.Remote);
			i++;
			row.createCell(i).setCellValue(listData.size() > 0 ? listData.get(0) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 1 ? listData.get(1) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 2 ? listData.get(2) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 3 ? listData.get(3) : "");

			listData = (List) d.get(Constant.FCC);
			i++;
			row.createCell(i).setCellValue(listData.size() > 0 ? listData.get(0) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 1 ? listData.get(1) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 2 ? listData.get(2) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 3 ? listData.get(3) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 4 ? listData.get(4) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 5 ? listData.get(5) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 6 ? listData.get(6) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 7 ? listData.get(7) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 8 ? listData.get(8) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() >9 ? listData.get(9) : "");

			listData = (List) d.get(Constant.Frequency);
			i++;
			row.createCell(i).setCellValue(listData.size() > 0 ? listData.get(0) : "");
			i++;
			row.createCell(i).setCellValue(listData.size() > 1 ? listData.get(1) : "");

			System.out.println("$$$$$$$" + d.get(Constant.Model).toString());
		
			listData = (List) d.get(Constant.newCol);
			for (String ss : listData) {
				String s= ss.toUpperCase();
				if (s.contains(Constant.CODE_SERIES) | s.contains(Constant.FCC) | s.contains("MHz")
						| s.contains(Constant.ILCO) | s.contains(Constant.JET) | s.contains(Constant.JMA)
						| s.contains(Constant.KEYLINE) | s.contains(Constant.LISHI21) | s.contains("LISH121")
						| s.contains(Constant.MASTER_TRANSPONDER_DATA) | s.contains("KEYS")
						| s.contains(Constant.MECHANICAL_KEY) | s.contains(Constant.OEM)
						| s.contains(Constant.PAGE_2_DATA) | s.contains(Constant.STRATTEC)
						| s.contains(Constant.TRANSPONDER) | s.contains(Constant.TRANSPONDER_RE_USE)
						| s.contains(Constant.VALET_TRANSPONDER_DATA) | s.contains("REMOTE") | s.contains("IC") | s.contains(Constant.PCB)| s.contains("ASSY")) {
					System.out.println(s);
				}
			}

		}

		// Resize all columns to fit the content size
		for (int i = 0; i < columns.length; i++) {
			sheet.autoSizeColumn(i);
		}

		// Write the output to a file
		FileOutputStream fileOut = null;
		try {
			fileOut = new FileOutputStream("D:\\Freelancer\\thoan_excel\\Results\\"+fileName);
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

	public Map<String, Object> createMap() {
		Map<String, Object> result = new HashMap<>();
		result.put(Constant.Page_Number, "");
		result.put(Constant.Year_From, "");
		result.put(Constant.Year_To, "");
		result.put(Constant.Make, "");
		result.put(Constant.Model, "");
		result.put(Constant.System_Type, "");
		// result.put(Constant.A_TEK, "");
		// result.put(Constant.BARNES, new ArrayList<>());
		// result.put(Constant.HATA, new ArrayList<>());
		// result.put(Constant.HILLMAN, "");
		// result.put(Constant.HOWARD_KEYS, "");
		// result.put(Constant.HYKO, "");
		result.put(Constant.EMERGENCY_KEY, new ArrayList<>());
		result.put(Constant.PROX, new ArrayList<>());
		result.put(Constant.EKEY, new ArrayList<>());
		result.put(Constant.ILCO, new ArrayList<>());
		result.put(Constant.JET, new ArrayList<>());
		result.put(Constant.JMA, new ArrayList<>());
		result.put(Constant.KEYLINE, new ArrayList<>());
		// result.put(Constant.LOCKCRAFT, new ArrayList<>());
		result.put(Constant.OEM, new ArrayList<>());
		result.put(Constant.PCB, new ArrayList<>());
		result.put(Constant.IC, new ArrayList<>());
		result.put(Constant.STRATTEC, new ArrayList<>());
		// result.put(Constant.WINZER, new ArrayList<>());
		// result.put(Constant.ACCU, "");
		// result.put(Constant.ACCU_READER, "");
		// result.put(Constant.BTR, "");
		// result.put(Constant.EEZ_READER, "");
		// result.put(Constant.KOBRA, "");
		result.put(Constant.LISHI, new ArrayList<>());
		result.put(Constant.LISHI21, new ArrayList<>());
		result.put(Constant.MECHANICAL_KEY, new ArrayList<>());
		result.put(Constant.CODE_SERIES, new ArrayList<>());
		result.put(Constant.MAX_KEYS, new ArrayList<>());
		result.put(Constant.TRANSPONDER, new ArrayList<>());
		result.put(Constant.MASTER_TRANSPONDER_DATA, new ArrayList<>());
		result.put(Constant.VALET_TRANSPONDER_DATA, "");
		result.put(Constant.PAGE_2_DATA, "");
		result.put(Constant.TRANSPONDER_RE_USE, "");
		result.put(Constant.Remote, new ArrayList<>());
		result.put(Constant.FCC, new ArrayList<>());
		result.put(Constant.Frequency, new ArrayList<>());
		result.put(Constant.newCol, new ArrayList<>());

		return result;
	}

	public static void main(String s[]) {
		JFrame frame = new JFrame("");
		ExcelProcessing6 panel = new ExcelProcessing6();
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