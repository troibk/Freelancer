package thoan_excel;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.nodes.Document;

import org.apache.poi.ss.usermodel.*;
import java.awt.event.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.awt.*;
import java.util.List;
import java.util.*;

public class AndyEastoe extends JPanel implements ActionListener {
	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;
	String path = "C:\\Users\\ha.vuthanh\\Desktop\\";
	JButton go;
	JFileChooser chooser;
	String choosertitle;
	File fileChoosen;
	private static List<String[]> datalist = new ArrayList<>();

	public AndyEastoe() {
		go = new JButton("Do it");
		go.addActionListener(this);
		add(go);
	}

	public void actionPerformed(ActionEvent e) {
		chooser = new JFileChooser();
		chooser.setCurrentDirectory(new java.io.File("."));
		chooser.setDialogTitle(choosertitle);
		chooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
		//
		// disable the "All files" option.
		//
		chooser.setAcceptAllFileFilterUsed(false);
		chooser.addChoosableFileFilter(new FileNameExtensionFilter("*.xlsx", "xlsx"));
		//
		if (chooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
			System.out.println("getCurrentDirectory(): " + chooser.getCurrentDirectory());
			System.out.println("getSelectedFile() : " + chooser.getSelectedFile());
			fileChoosen = chooser.getSelectedFile();
			processExcel2(fileChoosen);
		} else {
			System.out.println("No Selection ");
		}

	}

	public Dimension getPreferredSize() {
		return new Dimension(100, 100);
	}

	public void processExcel2(File file) {
		Document doc = null;
		try {
			FileInputStream inputFile = new FileInputStream(file);
			XSSFWorkbook wb = new XSSFWorkbook(inputFile);
			XSSFSheet sheet = wb.getSheetAt(0);

			if (sheet == null) {
				System.out.println("KO CO SHEET :" + sheet);
			} else {
				Map<String, Integer> columnNameMap = new HashMap<>();
				Row row = sheet.getRow(0);
				short minColIx = row.getFirstCellNum();
				short maxColIx = row.getLastCellNum();
				for (short colIx = minColIx; colIx < maxColIx; colIx++) {
					Cell cell = row.getCell(colIx);
					columnNameMap.put(cell.getStringCellValue().trim(), cell.getColumnIndex());
				}

				for (int i = 5850; i < sheet.getPhysicalNumberOfRows(); i++) {
					String isError = "false";
					String phone1 = "";
					String phone2 = "";
					String phone3 = "";
					String email = "";
					String madid = "";
					String fn = "";
					String ln = "";
					String zip = "";
					String ct = "";
					String st = "";
					String country = "";
					String dob = "";
					String doby = "";
					String gen = "";
					String age = "";
					String uid = "";
					String value = "";

					try {
						Cell First_name = sheet.getRow(i).getCell(columnNameMap.get("First name"));
						if (First_name != null && First_name.getCellType() == Cell.CELL_TYPE_STRING) {
							fn = First_name.getStringCellValue();
						} 

						Cell Last_name = sheet.getRow(i).getCell(columnNameMap.get("Last name"));
						if (Last_name != null && Last_name.getCellType() == Cell.CELL_TYPE_STRING) {
							ln = Last_name.getStringCellValue();
						} 

						Cell ID = sheet.getRow(i).getCell(columnNameMap.get("ID"));
						if (ID != null && ID.getCellType() == Cell.CELL_TYPE_STRING) {
							uid = ID.getStringCellValue();
						} 

						Cell City = sheet.getRow(i).getCell(columnNameMap.get("City"));
						if (City != null && City.getCellType() == Cell.CELL_TYPE_STRING) {
							ct = City.getStringCellValue();
						} 

						Cell State = sheet.getRow(i).getCell(columnNameMap.get("State"));
						if (State != null && State.getCellType() == Cell.CELL_TYPE_STRING) {
							st = State.getStringCellValue();
						} 

						Cell Country = sheet.getRow(i).getCell(columnNameMap.get("Country"));
						if (Country != null && Country.getCellType() == Cell.CELL_TYPE_STRING) {
							country = Country.getStringCellValue();
						} 

						Cell Postal_code = sheet.getRow(i).getCell(columnNameMap.get("Postal code"));
						if (Postal_code != null && Postal_code.getCellType() == Cell.CELL_TYPE_STRING) {
							zip = Postal_code.getStringCellValue();
						} 

						Cell Mobile_phone = sheet.getRow(i).getCell(columnNameMap.get("Mobile phone"));
						if (Mobile_phone != null && Mobile_phone.getCellType() == Cell.CELL_TYPE_STRING) {
							phone1 = Mobile_phone.getStringCellValue();
							String first = phone1.substring(0, 1);
							if (phone1.length() == 10 && first.equals("0")) {
								phone1 = "+61" + phone1.substring(1);
							} else if (phone1.length() == 9 && first.equals("4")) {
								phone1 = "+61" + phone1;
							} else if (phone1.length() == 9 && first.equals("4")) {
								phone1 = "+61" + phone1;
							} else if (phone1.length() >= 12 && !first.equals("+")) {
								phone1 = "+" + phone1;
							}else if (phone1.length()==11 && phone1.substring(0, 2).equals("61")) {
								phone1 = "+" + phone1;
							}
						} 

						Cell Home_phone = sheet.getRow(i).getCell(columnNameMap.get("Home phone"));
						if (Home_phone != null && Home_phone.getCellType() == Cell.CELL_TYPE_STRING) {
							phone2 = Home_phone.getStringCellValue();
							String first = phone2.substring(0, 1);
							if (phone2.length() == 10 && first.equals("0")) {
								phone2 = "+61" + phone2.substring(1);
							} else if (phone2.length() == 9 && first.equals("4")) {
								phone2 = "+61" + phone2;
							} else if (phone2.length() == 9 && first.equals("4")) {
								phone2 = "+61" + phone2;
							} else if (phone2.length() >= 12 && !first.equals("+")) {
								phone2 = "+" + phone2;
							}else if (phone2.length()==11 && phone2.substring(0, 2).equals("61")) {
								phone2 = "+" + phone2;
							}
						}

						Cell Work_phone = sheet.getRow(i).getCell(columnNameMap.get("Work phone"));
						if (Work_phone != null && Work_phone.getCellType() == Cell.CELL_TYPE_STRING) {
							phone3 = Work_phone.getStringCellValue();
							String first = phone3.substring(0, 1);
							if (phone3.length() == 10 && first.equals("0")) {
								phone3 = "+61" + phone3.substring(1);
							} else if (phone3.length() == 9 && first.equals("4")) {
								phone3 = "+61" + phone3;
							} else if (phone3.length() == 9 && first.equals("4")) {
								phone3 = "+61" + phone3;
							} else if (phone3.length() >= 12 && !first.equals("+")) {
								phone3 = "+" + phone3;
							}else if (phone3.length()==11 && phone3.substring(0, 2).equals("61")) {
								phone3 = "+" + phone3;
							}
						} 

						Cell Email = sheet.getRow(i).getCell(columnNameMap.get("Email"));
						if (Email != null && Email.getCellType() == Cell.CELL_TYPE_STRING) {
							email = Email.getStringCellValue();
						} 

						String[] data = new String[] { phone1, phone2, phone3, email, madid, fn, ln, zip, ct, st, country,
								dob, doby, gen, age, uid,value,"" };
						datalist.add(data);
						if (datalist.size() % 50 == 0) {
							System.out.println("======Write:" + datalist.size());
							appendExcelFile("Facebook_value_based_audience_file");
							Thread.sleep(1000);

						}

					} catch (Exception e) {
						System.out.println();
						String[] data = new String[] { phone1, phone2, phone3, email, madid, fn, ln, zip, ct, st, country,
								dob, doby, gen, age, uid,value, "false" };
						datalist.add(data);
					}
				}
				System.out.println("======Write:" + datalist.size());
				appendExcelFile("Facebook_value_based_audience_file");
				Thread.sleep(1000);
				System.out.println("======Finished=====");
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void processExcel1(File file) {
		Document doc = null;
		try {
			FileInputStream inputFile = new FileInputStream(file);
			XSSFWorkbook wb = new XSSFWorkbook(inputFile);
			XSSFSheet sheet = wb.getSheetAt(0);

			if (sheet == null) {
				System.out.println("KO CO SHEET :" + sheet);
			} else {
				Map<String, Integer> columnNameMap = new HashMap<>();
				Row row = sheet.getRow(0);
				short minColIx = row.getFirstCellNum();
				short maxColIx = row.getLastCellNum();
				for (short colIx = minColIx; colIx < maxColIx; colIx++) {
					Cell cell = row.getCell(colIx);
					columnNameMap.put(cell.getStringCellValue().trim(), cell.getColumnIndex());
				}

				Map<String, String> colourMaterialAbbr = new HashMap<String, String>();
				colourMaterialAbbr.put("B.", "Baby");
				colourMaterialAbbr.put("Blk", "Black");
				colourMaterialAbbr.put("Brn", "Brown");
				colourMaterialAbbr.put("Clr", "Clear");
				colourMaterialAbbr.put("D.", "Dark");
				colourMaterialAbbr.put("E.", "Emerald");
				colourMaterialAbbr.put("L.", "Light");
				colourMaterialAbbr.put("Met ", "Metallic ");
				colourMaterialAbbr.put("H.", "Hot");
				// colourMaterialAbbr.put("PL", "Pearl");
				colourMaterialAbbr.put("Slv", "Silver");
				// colourMaterialAbbr.put("TURQ", "Turquoise");
				colourMaterialAbbr.put("Wht", "White");
				colourMaterialAbbr.put("Camo", "Camouflage");
				// colourMaterialAbbr.put("FW", "Faux Wood");
				colourMaterialAbbr.put("Holo", "Hologram");
				colourMaterialAbbr.put("F.", "Faux");
				colourMaterialAbbr.put("M.", "Mini");
				colourMaterialAbbr.put("Gltr", "Glitter");
				// colourMaterialAbbr.put("LE", "Leather");
				// colourMaterialAbbr.put("M. FIBER", "MICROFIBER");
				colourMaterialAbbr.put("Str", "Stretch");
				// colourMaterialAbbr.put("VEL", "Velvet");
				colourMaterialAbbr.put("V. Suede", "Veggie Suede");
				colourMaterialAbbr.put("V. Leather", "Vegan Leather");
				colourMaterialAbbr.put("Pat", "Patent");
				// colourMaterialAbbr.put("PF", "Platform");
				// colourMaterialAbbr.put("MULTI GTR", "Multi Glitter");

				for (int i = 4500; i < sheet.getPhysicalNumberOfRows(); i++) {
					String[] data0 = new String[48];
					for (int j = 0; j < data0.length; j++) {
						data0[j] = "";
					}
					String isError = "false";
					try {
						row = sheet.getRow(i);
						if (row != null) {
							// Create Handle
							// PLEASER_ITEM
							Cell PLEASER_ITEM = sheet.getRow(i).getCell(columnNameMap.get("PLEASER_ITEM"));
							String PLEASER_ITEM_value = "";
							if (PLEASER_ITEM != null && PLEASER_ITEM.getCellType() == Cell.CELL_TYPE_STRING) {
								PLEASER_ITEM_value = PLEASER_ITEM.getStringCellValue();
							} else {
								// System.out.println("PLEASER_ITEM is null in row "+i);
								isError = "true";
							}
							// COLOUR_DESCRIPTION
							Cell COLOUR_DESCRIPTION = sheet.getRow(i).getCell(columnNameMap.get("COLOR_DESCRIPTION"));
							String COLOUR_DESCRIPTION_value = "";
							if (COLOUR_DESCRIPTION != null
									&& COLOUR_DESCRIPTION.getCellType() == Cell.CELL_TYPE_STRING) {
								COLOUR_DESCRIPTION_value = COLOUR_DESCRIPTION.getStringCellValue() + " ";
								for (Map.Entry<String, String> entry : colourMaterialAbbr.entrySet()) {
									if (!COLOUR_DESCRIPTION_value.contains(entry.getValue())) {
										COLOUR_DESCRIPTION_value = COLOUR_DESCRIPTION_value
												.replace(entry.getKey(), entry.getValue()).trim();
									}
								}

							} else {
								// System.out.println("COLOUR_DESCRIPTION is null in row "+i);
								isError = "true";
							}

							String handle = PLEASER_ITEM_value + "-" + COLOUR_DESCRIPTION_value; // Create Handle
							handle = handle.replaceAll("/", "-");
							handle = handle.replaceAll(" ", "-");
							handle = handle.replaceAll("\\.", "");
							handle = handle.replaceAll("\\)", "-");
							handle = handle.replaceAll("\\(", "-");
							handle = handle.replaceAll("--", "-");
							handle = handle.toLowerCase();
							// Create Title
							// STYLE_NAME
							Cell STYLE_NAME = sheet.getRow(i).getCell(columnNameMap.get("STYLE_NAME"));
							String STYLE_NAME_value = "";
							if (STYLE_NAME != null && STYLE_NAME.getCellType() == Cell.CELL_TYPE_STRING) {
								STYLE_NAME_value = STYLE_NAME.getStringCellValue();
							} else {
								// System.out.println("STYLE_NAME is null in row "+i);
								isError = "true";
							}
							// HEEL_HEIGHT_IN_INCH
							Cell HEEL_HEIGHT_IN_INCH = sheet.getRow(i)
									.getCell(columnNameMap.get("HEEL_HEIGHT_IN_INCH"));
							double HEEL_HEIGHT_IN_INCH_value = 0;
							if (HEEL_HEIGHT_IN_INCH != null
									&& HEEL_HEIGHT_IN_INCH.getCellType() == Cell.CELL_TYPE_NUMERIC) {
								HEEL_HEIGHT_IN_INCH_value = HEEL_HEIGHT_IN_INCH.getNumericCellValue();
							} else {
								// System.out.println("HEEL_HEIGHT_IN_INCH is null in row "+i);
								isError = "true";
							}
							// 2ND_SUB_CATEGORY
							String THIRD_SUB_CATEGORY_value = "";
							Cell SEC_SUB_CATEGORY = sheet.getRow(i).getCell(columnNameMap.get("2ND_SUB_CATEGORY"));
							String SEC_SUB_CATEGORY_value = "";
							if (SEC_SUB_CATEGORY == null) {
								THIRD_SUB_CATEGORY_value = "Platform";
							} else if (SEC_SUB_CATEGORY.getCellType() == Cell.CELL_TYPE_STRING) {
								SEC_SUB_CATEGORY_value = SEC_SUB_CATEGORY.getStringCellValue().toUpperCase();
								if (!SEC_SUB_CATEGORY_value.contains("FLATS")
										&& !SEC_SUB_CATEGORY_value.contains("SWAN")
										&& !SEC_SUB_CATEGORY_value.contains("PRESTIGE")
										&& !SEC_SUB_CATEGORY_value.contains("GORGEOUS")
										&& !SEC_SUB_CATEGORY_value.contains("DELUXE")) {
									THIRD_SUB_CATEGORY_value = THIRD_SUB_CATEGORY_value + "Platform ";
								}

								if (SEC_SUB_CATEGORY_value.contains("HEEL")) {
									THIRD_SUB_CATEGORY_value = THIRD_SUB_CATEGORY_value + "Heel";
								}

								if (SEC_SUB_CATEGORY_value.contains("Ankle/Mid-Calf Boots".toUpperCase())) {
									THIRD_SUB_CATEGORY_value = THIRD_SUB_CATEGORY_value + "Mid Calf Boot ";
								}

								if (SEC_SUB_CATEGORY_value.contains("Thigh High Boots".toUpperCase())) {
									THIRD_SUB_CATEGORY_value = THIRD_SUB_CATEGORY_value + "Thigh High Boot ";
								}

							} else {
								// System.out.println("SEC_SUB_CATEGORY is null in row "+i);
								isError = "true";
							}

							String title = "";
							COLOUR_DESCRIPTION_value = COLOUR_DESCRIPTION_value.replace("LE.", "Leather");
							COLOUR_DESCRIPTION_value = COLOUR_DESCRIPTION_value.replace("RSTN", "Rhinestone");
							COLOUR_DESCRIPTION_value = COLOUR_DESCRIPTION_value.replace("UV", "Ultraviolet");
							if (HEEL_HEIGHT_IN_INCH_value > 0) {
								title = (STYLE_NAME_value + " | " + HEEL_HEIGHT_IN_INCH_value + " INCH  "
										+ COLOUR_DESCRIPTION_value + " " + THIRD_SUB_CATEGORY_value).trim()
												.toUpperCase(); // Create Title
								title = title.replace(".0", "");
							} else {
								title = (STYLE_NAME_value + " | " + COLOUR_DESCRIPTION_value + " "
										+ THIRD_SUB_CATEGORY_value).trim().toUpperCase(); // Create Title
							}

							// Create Body (HTML)
							// DESCRIPTION_ADDITION
							Cell DESCRIPTION_ADDITION = sheet.getRow(i)
									.getCell(columnNameMap.get("DESCRIPTION_ADDITION"));
							String DESCRIPTION_ADDITION_value = "";
							if (DESCRIPTION_ADDITION != null
									&& DESCRIPTION_ADDITION.getCellType() == Cell.CELL_TYPE_STRING) {
								DESCRIPTION_ADDITION_value = DESCRIPTION_ADDITION.getStringCellValue();
							} else {
								// System.out.println("DESCRIPTION_ADDITION is null in row "+i);
								// isError="true";
							}

							Cell DESCRIPTION = sheet.getRow(i).getCell(columnNameMap.get("DESCRIPTION"));
							String DESCRIPTION_value = "";
							if (DESCRIPTION != null && DESCRIPTION.getCellType() == Cell.CELL_TYPE_STRING) {
								DESCRIPTION_value = DESCRIPTION.getStringCellValue();
							} else {
								// System.out.println("DESCRIPTION is null in row "+i);
								// isError="true";
							}

							if (DESCRIPTION_ADDITION_value.isEmpty()) {
								DESCRIPTION_ADDITION_value = DESCRIPTION_value;
							}

							// UPPER_MATERIAL
							Cell UPPER_MATERIAL = sheet.getRow(i).getCell(columnNameMap.get("UPPER_MATERIAL"));
							String UPPER_MATERIAL_value = "";
							if (UPPER_MATERIAL != null && UPPER_MATERIAL.getCellType() == Cell.CELL_TYPE_STRING) {
								UPPER_MATERIAL_value = UPPER_MATERIAL.getStringCellValue();
							} else {
								// System.out.println("UPPER_MATERIAL is null in row "+i);
								// isError="true";
							}

							// MID-SOLE_MATERIAL
							Cell MID_SOLE_MATERIAL = sheet.getRow(i).getCell(columnNameMap.get("MID-SOLE_MATERIAL"));
							String MID_SOLE_MATERIAL_value = "";
							if (MID_SOLE_MATERIAL != null && MID_SOLE_MATERIAL.getCellType() == Cell.CELL_TYPE_STRING) {
								MID_SOLE_MATERIAL_value = MID_SOLE_MATERIAL.getStringCellValue();
							} else {
								// System.out.println("MID-SOLE_MATERIAL is null in row "+i);
								// isError="true";
							}

							// LINING_MATERIAL
							Cell LINING_MATERIAL = sheet.getRow(i).getCell(columnNameMap.get("LINING_MATERIAL"));
							String LINING_MATERIAL_value = "";
							if (LINING_MATERIAL != null && LINING_MATERIAL.getCellType() == Cell.CELL_TYPE_STRING) {
								LINING_MATERIAL_value = LINING_MATERIAL.getStringCellValue();
							} else {
								// System.out.println("LINING_MATERIAL is null in row "+i);
								// isError="true";
							}

							// OUTSOLE_MATERIAL
							Cell OUTSOLE_MATERIAL = sheet.getRow(i).getCell(columnNameMap.get("OUTSOLE_MATERIAL"));
							String OUTSOLE_MATERIAL_value = "";
							if (OUTSOLE_MATERIAL != null && OUTSOLE_MATERIAL.getCellType() == Cell.CELL_TYPE_STRING) {
								OUTSOLE_MATERIAL_value = OUTSOLE_MATERIAL.getStringCellValue();
							} else {
								// System.out.println("OUTSOLE_MATERIAL is null in row "+i);
								// isError="true";
							}

							// VEGAN_STYLE
							Cell VEGAN_STYLE = sheet.getRow(i).getCell(columnNameMap.get("VEGAN_STYLE"));
							String VEGAN_STYLE_value = "";
							if (VEGAN_STYLE != null && VEGAN_STYLE.getCellType() == Cell.CELL_TYPE_STRING) {
								VEGAN_STYLE_value = VEGAN_STYLE.getStringCellValue();
							} else {
								// System.out.println("VEGAN_STYLE is null in row "+i);
								isError = "true";
							}

							// FIT_GUIDE
							Cell FIT_GUIDE = sheet.getRow(i).getCell(columnNameMap.get("FIT_GUIDE"));
							String FIT_GUIDE_value = "";
							if (FIT_GUIDE != null && FIT_GUIDE.getCellType() == Cell.CELL_TYPE_STRING) {
								FIT_GUIDE_value = FIT_GUIDE.getStringCellValue();
							} else {
								// System.out.println("FIT_GUIDE is null in row "+i);
								isError = "true";
							}

							String html = "<strong> " + title + "</strong>\r\n" + "<div><br></div>\r\n"
									+ "<div><meta charset=\"utf-8\">\r\n" + DESCRIPTION_ADDITION_value + "</div>\r\n"
									+ "<meta charset=\"utf-8\"> <b>Item # (SKU):</b> \r\n" + PLEASER_ITEM_value
									+ "<!-- TABS -->\r\n" + "<h5>Product Materials</h5>\r\n" + "<div>\r\n" + "<li> \r\n"
									+ COLOUR_DESCRIPTION_value + "</li>\r\n" + "</div>\r\n" + "<div>\r\n"
									+ "<li>Upper Material:\r\n" + UPPER_MATERIAL_value + "</li>\r\n"
									+ "<li>Mid Sole Material:\r\n" + MID_SOLE_MATERIAL_value + "</li>\r\n"
									+ "<li>Lining Material:\r\n" + LINING_MATERIAL_value + "</li>\r\n"
									+ "<li>Outsole Material: \r\n" + OUTSOLE_MATERIAL_value + "<br></li>\r\n";
							if (VEGAN_STYLE_value.toUpperCase().equals("YES")) {
								html = html + "<li>Vegan Style: Yes</li>";
							}

							html = html + "</div>\r\n" + "<h5>Sizing Guide &amp; Tips<br></h5>\r\n" + "<ul>\r\n"
									+ "<li><span style=\"color: #000000;\" data-mce-style=\"color: #000000;\"><span data-mce-fragment=\"1\">The&nbsp;<meta charset=\"utf-8\">  "
									+ PLEASER_ITEM_value + " fit is " + FIT_GUIDE_value + ".</span></span></li>\r\n"
									+ "</ul>\r\n"
									+ "<p><span style=\"color: #000000;\" data-mce-style=\"color: #000000;\"><span data-mce-fragment=\"1\">Please refer to the US sizing when selecting your pair of heels as this is closest to AU sizing. If ordering this style with a closed toe or are in between sizing we generally advise that you size up.</span> If you have any other questions about sizing be sure to use the chat box where you'll&nbsp;get expert advice from a team who live and breathe pole dance!<br><br></span><img style=\"color: #000000; font-family: -apple-system, BlinkMacSystemFont, 'San Francisco', 'Segoe UI', Roboto, 'Helvetica Neue', sans-serif; font-size: 1.4em;\" src=\"https://cdn.shopify.com/s/files/1/0271/4107/9155/files/size-guide_480x480.jpg?v=1607908259\" alt=\"\" data-mce-fragment=\"1\" data-mce-src=\"https://cdn.shopify.com/s/files/1/0271/4107/9155/files/size-guide_480x480.jpg?v=1607908259\" data-mce-style=\"color: #000000; font-family: -apple-system, BlinkMacSystemFont, 'San Francisco', 'Segoe UI', Roboto, 'Helvetica Neue', sans-serif; font-size: 1.4em;\"></p>\r\n"
									+ "<h5>About Pleaser Heels</h5>\r\n"
									+ "<p><span style=\"color: #000000;\" data-mce-style=\"color: #000000;\">Get the best iconic footwear that's fun, evocative, and vivaciously genuine in its appeal. When you buy Pleaser shoes from The Pole Room, you're getting heels that are made with a<span data-mce-fragment=\"1\">ttention to detail, superb craftsmanship, fine materials, innovative designs, unmatched selection, and unbeatable prices. All of these have contributed to PLEASERs unparalleled success and reputation among sexy shoe aficionados and professional performers alike. Today, PLEASER is undoubtedly one of the most renowned sexy shoe brands worldwide!</span> They are perfect to increase the grace of every pole lover wardrobe.</span></p>\r\n"
									+ "<h5><span style=\"color: #000000;\" data-mce-style=\"color: #000000;\">Shipping Information</span></h5>\r\n"
									+ "<p><span style=\"color: #000000;\" data-mce-style=\"color: #000000;\"><span>We hold stock in our Melbourne warehouse that is&nbsp;dispatched to our customers within 2 business days. In the event that we do not have stock at our Melbourne warehouse there may be an option to purchase from our US warehouse. Dispatch times from the US is approx 5 business days to our Melbourne warehouse and your order will then be shipped together with any other items purchased.</span></span></p>\r\n"
									+ "<h5>Returns&nbsp;Guarantee</h5>\r\n"
									+ "<p><span style=\"color: #000000;\" data-mce-style=\"color: #000000;\">We want you to absolutely love your new purchase! If you are not happy for any reason then we'll make it easy to return your item to us.&nbsp;</span></p>\r\n"
									+ "<ul>\r\n"
									+ "<li><span style=\"color: #000000;\" data-mce-style=\"color: #000000;\">Clearance or discontinued stock cannot be returned or exchanged. </span></li>\r\n"
									+ "<li><span style=\"color: #000000;\" data-mce-style=\"color: #000000;\">We only replace items if they are defective or damaged.</span></li>\r\n"
									+ "<li><span style=\"color: #000000;\" data-mce-style=\"color: #000000;\">Our return and exchange guarantee lasts 30 days.</span><br></li>\r\n"
									+ "<li><span style=\"color: #000000;\" data-mce-style=\"color: #000000;\">If the item does not fit and you wish to return, you will pay postage costs and a credit note will be provided upon the item being returned (pending inspection).</span></li>\r\n"
									+ "</ul>\r\n"
									+ "<p><span style=\"color: #000000;\" data-mce-style=\"color: #000000;\">Our complete Returns &amp; Exchanges policy can be found <a href=\"https://shop.thepoleroom.com.au/policies/refund-policy\">here</a>.&nbsp;</span></p>\r\n"
									+ "<!-- /TABS -->\r\n";

							// Create Vendor
							// BRAND
							Cell BRAND = sheet.getRow(i).getCell(columnNameMap.get("BRAND"));
							String BRAND_value = "";
							if (BRAND != null && BRAND.getCellType() == Cell.CELL_TYPE_STRING) {
								BRAND_value = BRAND.getStringCellValue();
							} else {
								// System.out.println("BRAND is null in row "+i);
								isError = "true";
							}
							String vendor = BRAND_value;

							// Create Type
							String type = "Pleaser Heels";

							// Create Tags
							// ITEM_STATUS
							Cell ITEM_STATUS = sheet.getRow(i).getCell(columnNameMap.get("ITEM_STATUS"));
							String ITEM_STATUS_value = "";
							if (ITEM_STATUS != null && ITEM_STATUS.getCellType() == Cell.CELL_TYPE_STRING) {
								ITEM_STATUS_value = ITEM_STATUS.getStringCellValue();
							} else {
								// System.out.println("ITEM_STATUS is null in row "+i);
								isError = "true";
							}
							String heighTag = "";
							if (HEEL_HEIGHT_IN_INCH_value < 1) {
								heighTag = "";
							} else if (HEEL_HEIGHT_IN_INCH_value < 2) {
								heighTag = "Height_" + "1 Inch";
							} else if (HEEL_HEIGHT_IN_INCH_value < 3) {
								heighTag = "Height_" + "2 Inch";
							} else if (HEEL_HEIGHT_IN_INCH_value < 4) {
								heighTag = "Height_" + "3 Inch";
							} else if (HEEL_HEIGHT_IN_INCH_value < 5) {
								heighTag = "Height_" + "4 Inch";
							} else if (HEEL_HEIGHT_IN_INCH_value < 6) {
								heighTag = "Height_" + "5 Inch";
							} else if (HEEL_HEIGHT_IN_INCH_value < 7) {
								heighTag = "Height_" + "6 Inch";
							} else if (HEEL_HEIGHT_IN_INCH_value < 8) {
								heighTag = "Height_" + "7 Inch";
							} else if (HEEL_HEIGHT_IN_INCH_value < 9) {
								heighTag = "Height_" + "8 Inch";
							} else if (HEEL_HEIGHT_IN_INCH_value < 10) {
								heighTag = "Height_" + "9 Inch";
							} else if (HEEL_HEIGHT_IN_INCH_value < 11) {
								heighTag = "Height_" + "10 Inch";
							} else if (HEEL_HEIGHT_IN_INCH_value < 12) {
								heighTag = "Height_" + "11 Inch";
							} else if (HEEL_HEIGHT_IN_INCH_value < 14) {
								heighTag = "Height_" + "13 Inch";
							} else if (HEEL_HEIGHT_IN_INCH_value == 20) {
								heighTag = "Height_" + "20 Inch";
							}

							String statusTag = "";
							if (ITEM_STATUS_value.equals("CURRENT")) {
								statusTag = "Status_" + "Current";
							} else if (ITEM_STATUS_value.equals("COMINGSOON")) {
								statusTag = "Status_" + "Coming Soon";
							} else if (ITEM_STATUS_value.equals("NEW")) {
								statusTag = "Status_" + "New";
							} else if (ITEM_STATUS_value.equals("PREORDER")) {
								statusTag = "Status_" + "Pre Order";
							} else if (ITEM_STATUS_value.equals("SALE")) {
								statusTag = "Status_" + "Sale";
							}

							String veganTag = "";
							if (VEGAN_STYLE_value.equals("YES")) {
								veganTag = "Vegan_Yes";
							}

							String styleTag = "";
							String[] styleFomats = STYLE_NAME_value.split("-");
							if (styleFomats.length > 0) {
								styleTag = styleFomats[0].toLowerCase();
								styleTag = styleTag.substring(0, 1).toUpperCase() + styleTag.substring(1);
								styleTag = "Style_" + styleTag;
							}

							String heelTag = "";
							DESCRIPTION_value = DESCRIPTION_value.toUpperCase();
							if (DESCRIPTION_value.contains("MID-CALF BOOT")) {
								heelTag = heelTag + "Heel_Mid-calf Boot,";
							}
							if (DESCRIPTION_value.contains("ANKLE BOOT")) {
								heelTag = heelTag + "Heel_Ankle Boot,";
							}
							if (DESCRIPTION_value.contains("SANDAL")) {
								heelTag = heelTag + "Heel_Sandal,";
							}
							if (DESCRIPTION_value.contains("PLATFORM")) {
								heelTag = heelTag + "Heel_Platform,";
							}
							if (DESCRIPTION_value.contains("THIGH BOOT")) {
								heelTag = heelTag + "Heel_Thigh Boot,";
							}
							if (DESCRIPTION_value.contains("KNEE BOOT")) {
								heelTag = heelTag + "Heel_Knee Boot,";
							}
							if (DESCRIPTION_value.contains("BOOT")) {
								heelTag = heelTag + "Heel_Boot,";
							}
							if (DESCRIPTION_value.contains("ANKLE STRAP SANDAL")) {
								heelTag = heelTag + "Heel_Ankle Starp Sandal,";
							}
							if (DESCRIPTION_value.contains("FRONT LACE UP")) {
								heelTag = heelTag + "Heel_Front Lace Up,";
							}
							if (DESCRIPTION_value.contains("SIDE LACE UP")) {
								heelTag = heelTag + "Heel_Side Lace Up,";
							}
							if (DESCRIPTION_value.contains("BOOTIE")) {
								heelTag = heelTag + "Heel_Bootie,";
							}
							if (DESCRIPTION_value.contains("STRAP")) {
								heelTag = heelTag + "Heel_Strap,";
							}

							if (!heelTag.isEmpty()) {
								heelTag = heelTag.substring(0, heelTag.lastIndexOf(","));
							}

							String materialTag = "";
							COLOUR_DESCRIPTION_value = COLOUR_DESCRIPTION_value.toUpperCase();
							if (COLOUR_DESCRIPTION_value.contains("GLITTER")) {
								materialTag = materialTag + "Material_Glitter,";
							}
							if (COLOUR_DESCRIPTION_value.contains("CHROME")) {
								materialTag = materialTag + "Material_Chrome,";
							}
							if (COLOUR_DESCRIPTION_value.contains("LEATHER")) {
								materialTag = materialTag + "Material_Leather,";
							}
							if (COLOUR_DESCRIPTION_value.contains("PATENT")) {
								materialTag = materialTag + "Material_Patent,";
							}
							if (COLOUR_DESCRIPTION_value.contains("SEQUINS")) {
								materialTag = materialTag + "Material_Sequins,";
							}
							if (COLOUR_DESCRIPTION_value.contains("HOLOGRAM")) {
								materialTag = materialTag + "Material_Hologram,";
							}
							if (COLOUR_DESCRIPTION_value.contains("RHINESTONE")) {
								materialTag = materialTag + "Material_Rhinestone,";
							}
							if (COLOUR_DESCRIPTION_value.contains("SUEDE")) {
								materialTag = materialTag + "Material_Suede,";
							}
							if (COLOUR_DESCRIPTION_value.contains("METALLIC")) {
								materialTag = materialTag + "Material_Metallic,";
							}
							if (COLOUR_DESCRIPTION_value.contains("TEXTURED")) {
								materialTag = materialTag + "Material_Textured,";
							}
							if (COLOUR_DESCRIPTION_value.contains("STRETCH")) {
								materialTag = materialTag + "Material_Stretch,";
							}
							if (COLOUR_DESCRIPTION_value.contains("VELVET")) {
								materialTag = materialTag + "Material_Velvet,";
							}
							if (COLOUR_DESCRIPTION_value.contains("ULTRAVIOLET")
									|| DESCRIPTION_value.contains("ILLUMINATOR") || DESCRIPTION_value.contains("NEON")
									|| DESCRIPTION_value.contains("REFL")) {
								materialTag = materialTag + "Material_Glow,";
							}
							if (!materialTag.isEmpty()) {
								materialTag = materialTag.substring(0, materialTag.lastIndexOf(","));
							}

							String designTag = "";
							COLOUR_DESCRIPTION_value = COLOUR_DESCRIPTION_value.toUpperCase();
							if (COLOUR_DESCRIPTION_value.contains("FADED")) {
								designTag = designTag + "Design_Faded,";
							}
							if (COLOUR_DESCRIPTION_value.contains("CLEAR")) {
								designTag = designTag + "Design_Clear,";
							}
							if (COLOUR_DESCRIPTION_value.contains("FROSTED")) {
								designTag = designTag + "Design_Frosted,";
							}
							if (COLOUR_DESCRIPTION_value.contains("MULTI")) {
								designTag = designTag + "Design_Multi,";
							}
							if (COLOUR_DESCRIPTION_value.contains("PATTERN")) {
								designTag = designTag + "Design_Pattern,";
							}
							if (COLOUR_DESCRIPTION_value.contains("PRINT")) {
								designTag = designTag + "Design_Print,";
							}
							if (COLOUR_DESCRIPTION_value.contains("RAINBOW")) {
								designTag = designTag + "Design_Rainbow,";
							}
							if (COLOUR_DESCRIPTION_value.contains("REFLECTIVE")) {
								designTag = designTag + "Design_Reflective,";
							}
							if (COLOUR_DESCRIPTION_value.contains("UNIVERSE")) {
								designTag = designTag + "Design_Universe,";
							}
							if (COLOUR_DESCRIPTION_value.contains("CAMOUFLAGE")) {
								designTag = designTag + "Design_Camouflage,";
							}
							if (COLOUR_DESCRIPTION_value.contains("TINTED")) {
								designTag = designTag + "Design_Tinted,";
							}
							if (COLOUR_DESCRIPTION_value.contains("SHAPE")) {
								designTag = designTag + "Design_Shape,";
							}
							if (!designTag.isEmpty()) {
								designTag = designTag.substring(0, designTag.lastIndexOf(","));
							}

							String tags = "";
							if (!heighTag.isEmpty()) {
								tags = tags + heighTag + ", ";
							}

							if (!statusTag.isEmpty()) {
								tags = tags + statusTag + ", ";
							}

							if (!styleTag.isEmpty()) {
								tags = tags + styleTag + ", ";
							}

							if (!veganTag.isEmpty()) {
								tags = tags + veganTag + ", ";
							}

							if (!designTag.isEmpty()) {
								tags = tags + designTag + ", ";
							}

							if (!materialTag.isEmpty()) {
								tags = tags + materialTag + ", ";
							}

							if (!heelTag.isEmpty()) {
								tags = tags + heelTag + ", ";
							}

							tags = tags.substring(0, tags.lastIndexOf(","));

							// Create Published
							String published = "TRUE";

							// Create Option 1 Name
							String option1Name = "Color";

							// Create Option 1 Value
							String option1Value = "White";
							if (COLOUR_DESCRIPTION_value.toUpperCase().contains("BLACK")) {
								option1Value = "Black";
							} else if (COLOUR_DESCRIPTION_value.toUpperCase().contains("PINK")) {
								option1Value = "Pink";
							} else if (COLOUR_DESCRIPTION_value.toUpperCase().contains("RED")) {
								option1Value = "Red";
							} else if (COLOUR_DESCRIPTION_value.toUpperCase().contains("PURPLE")) {
								option1Value = "Purple";
							} else if (COLOUR_DESCRIPTION_value.toUpperCase().contains("BLUE")) {
								option1Value = "Blue";
							} else if (COLOUR_DESCRIPTION_value.toUpperCase().contains("BLACK")) {
								option1Value = "Black";
							} else if (COLOUR_DESCRIPTION_value.toUpperCase().contains("GREEN")) {
								option1Value = "Green";
							} else if (COLOUR_DESCRIPTION_value.toUpperCase().contains("BROWN")) {
								option1Value = "Brown";
							} else if (COLOUR_DESCRIPTION_value.toUpperCase().contains("GREY")) {
								option1Value = "Grey";
							} else if (COLOUR_DESCRIPTION_value.toUpperCase().contains("GOLD")) {
								option1Value = "Yellow";
							} else if (COLOUR_DESCRIPTION_value.toUpperCase().contains("YELLOW")) {
								option1Value = "Yellow";
							} else if (COLOUR_DESCRIPTION_value.toUpperCase().contains("VIOLET")) {
								option1Value = "Purple";
							} else if (COLOUR_DESCRIPTION_value.toUpperCase().contains("TURQUOISE")) {
								option1Value = "Green";
							} else if (COLOUR_DESCRIPTION_value.toUpperCase().contains("TAN")) {
								option1Value = "Green";
							} else if (COLOUR_DESCRIPTION_value.toUpperCase().contains("RUBY")) {
								option1Value = "Pink";
							} else if (COLOUR_DESCRIPTION_value.toUpperCase().contains("ROSE")) {
								option1Value = "Red";
							} else if (COLOUR_DESCRIPTION_value.toUpperCase().contains("ORANGE")) {
								option1Value = "Orange";
							} else if (COLOUR_DESCRIPTION_value.toUpperCase().contains("LAVENDER")) {
								option1Value = "Purple";
							} else if (COLOUR_DESCRIPTION_value.toUpperCase().contains("RASBERRY")) {
								option1Value = "Red";
							}

							// Create Option 2 Name
							String option2Name = "Size";

							// Create Option 2 Value
							// SIZE_RANGE
							Cell SIZE_RANGE = sheet.getRow(i).getCell(columnNameMap.get("SIZE_RANGE"));
							String SIZE_RANGE_value = "";
							String[] option2Value = null;
							if (SIZE_RANGE != null) {
								if (SIZE_RANGE.getCellType() == Cell.CELL_TYPE_STRING) {
									SIZE_RANGE_value = SIZE_RANGE.getStringCellValue();
									if (SIZE_RANGE_value.equals("ONE SIZE")) {
										option2Value = new String[] { "US 1" };
									} else if (SIZE_RANGE_value.equals("S, M, L, XL")) {
										option2Value = SIZE_RANGE_value.split(",");
									} else if (SIZE_RANGE_value.startsWith("Men")) {
										SIZE_RANGE_value = SIZE_RANGE_value
												.substring(SIZE_RANGE_value.lastIndexOf(" "));
										int first = Integer.parseInt(SIZE_RANGE_value.split("-")[0].trim());
										int second = Integer.parseInt(SIZE_RANGE_value.split("-")[1].trim());
										if (first > second) {
											int temp = first;
											first = second;
											second = temp;
										}
										option2Value = new String[second - first + 1];
										int start = 0;
										for (int k = first; k <= second; k++) {
											option2Value[start] = "US " + k;
											start = start + 1;
										}
									}
								} else if (SIZE_RANGE.getCellType() == Cell.CELL_TYPE_NUMERIC) {
									int number = (int) SIZE_RANGE.getNumericCellValue();
									if (number < 20) {
										SIZE_RANGE_value = "US " + number;
										option2Value = new String[] { SIZE_RANGE_value };
									} else {
										Date dateTemp = SIZE_RANGE.getDateCellValue();
										Calendar cal = Calendar.getInstance();
										cal.setTime(dateTemp);
										int first = cal.get(Calendar.MONTH) + 1;
										int second = cal.get(Calendar.DAY_OF_MONTH);
										if (first > second) {
											int temp = first;
											first = second;
											second = temp;
										}
										option2Value = new String[second - first + 1];
										int start = 0;
										for (int k = first; k <= second; k++) {
											option2Value[start] = "US " + k;
											start = start + 1;
										}
									}
								} else if (SIZE_RANGE.getCellType() == Cell.CELL_TYPE_BLANK) {
									option2Value = new String[] { "" };
								}
							} else {
								// System.out.println("SIZE_RANGE is null in row "+i);
								option2Value = new String[] { "US 1" };
								isError = "true";
							}

							// Create Variant SKU
							String variantSKU = PLEASER_ITEM_value;

							// Create Variant Grams
							String variantGrams = "1200";

							// Create Variant Inventory Tracker
							String variantInventoryTracker = "shopify";

							// Create Variant Inventory Policy
							String variantInventoryPolicy = "deny";

							// Create Variant Fulfillment Service
							String variantFulfillmentService = "manual";

							// Create Variant Price
							Cell MSRP_AUD = sheet.getRow(i).getCell(columnNameMap.get("MSRP_AUD"));
							double MSRP_AUD_value = 0;
							if (MSRP_AUD != null && MSRP_AUD.getCellType() == Cell.CELL_TYPE_NUMERIC) {
								MSRP_AUD_value = MSRP_AUD.getNumericCellValue();
							} else {
								// System.out.println("MSRP_AUD is null in row "+i);
								isError = "true";
							}

							String variantCompareAtPrice = MSRP_AUD_value == 0 ? "" : ("" + MSRP_AUD_value);

							// VARIANT_COMPARE_AT_PRICE
							double variantPriceNumber = 0;
							if (ITEM_STATUS_value.equals("CURRENT")) {
								variantPriceNumber = round(MSRP_AUD_value * 93 / 100, 2);
							} else if (ITEM_STATUS_value.equals("COMINGSOON")) {
								variantPriceNumber = round(MSRP_AUD_value * 97 / 100, 2);
							} else if (ITEM_STATUS_value.equals("NEW")) {
								variantPriceNumber = round(MSRP_AUD_value * 95 / 100, 2);
							} else if (ITEM_STATUS_value.equals("PREORDER")) {
								variantPriceNumber = round(MSRP_AUD_value * 97 / 100, 2);
							} else if (ITEM_STATUS_value.equals("SALE")) {
								variantPriceNumber = round(MSRP_AUD_value * 90 / 100, 2);
							}

							String variantPrice = variantPriceNumber == 0 ? "" : ("" + variantPriceNumber);

							// Create Variant Requires Shipping
							String variantRequiresShipping = "TRUE";

							// Create Variant Taxable
							String variantTaxable = "TRUE";

							// Create Image Src 0
							Cell IMAGE_FULL = sheet.getRow(i).getCell(columnNameMap.get("IMAGE_FULL"));
							String IMAGE_FULL_value = "";
							if (IMAGE_FULL != null && IMAGE_FULL.getCellType() == Cell.CELL_TYPE_STRING) {
								IMAGE_FULL_value = IMAGE_FULL.getStringCellValue();
							} else {
								System.out.println("IMAGE_FULL is null in row " + i);
								isError = "true";
							}

							// Create Image Src 1
							Cell MULTIVIEW_IMAGE_1 = sheet.getRow(i).getCell(columnNameMap.get("MULTIVIEW_IMAGE_1"));
							String MULTIVIEW_IMAGE_1_value = "";
							if (MULTIVIEW_IMAGE_1 != null && MULTIVIEW_IMAGE_1.getCellType() == Cell.CELL_TYPE_STRING) {
								MULTIVIEW_IMAGE_1_value = MULTIVIEW_IMAGE_1.getStringCellValue();
							} else {
								// System.out.println("MULTIVIEW_IMAGE_1 is null in row "+i);
								isError = "true";
							}

							// Create Image Src 2
							Cell MULTIVIEW_IMAGE_2 = sheet.getRow(i).getCell(columnNameMap.get("MULTIVIEW_IMAGE_2"));
							String MULTIVIEW_IMAGE_2_value = "";
							if (MULTIVIEW_IMAGE_2 != null && MULTIVIEW_IMAGE_2.getCellType() == Cell.CELL_TYPE_STRING) {
								MULTIVIEW_IMAGE_2_value = MULTIVIEW_IMAGE_2.getStringCellValue();
							} else {
								// System.out.println("MULTIVIEW_IMAGE_2 is null in row "+i);
								isError = "true";
							}

							// Create Image Src 3
							Cell MULTIVIEW_IMAGE_3 = sheet.getRow(i).getCell(columnNameMap.get("MULTIVIEW_IMAGE_3"));
							String MULTIVIEW_IMAGE_3_value = "";
							if (MULTIVIEW_IMAGE_3 != null && MULTIVIEW_IMAGE_3.getCellType() == Cell.CELL_TYPE_STRING) {
								MULTIVIEW_IMAGE_3_value = MULTIVIEW_IMAGE_3.getStringCellValue();
							} else {
								// System.out.println("MULTIVIEW_IMAGE_3 is null in row "+i);
								isError = "true";
							}

							// Create Image Src 4
							Cell MULTIVIEW_IMAGE_4 = sheet.getRow(i).getCell(columnNameMap.get("MULTIVIEW_IMAGE_4"));
							String MULTIVIEW_IMAGE_4_value = "";
							if (MULTIVIEW_IMAGE_4 != null && MULTIVIEW_IMAGE_4.getCellType() == Cell.CELL_TYPE_STRING) {
								MULTIVIEW_IMAGE_4_value = MULTIVIEW_IMAGE_4.getStringCellValue();
							} else {
								// System.out.println("MULTIVIEW_IMAGE_4 is null in row "+i);
								isError = "true";
							}

							// Create Image Position - Numbers
							Map<Integer, String> multiviewImagesMap = new HashMap<Integer, String>();

							if (!MULTIVIEW_IMAGE_1_value.isEmpty()) {
								multiviewImagesMap.put(2, MULTIVIEW_IMAGE_1_value);
							}

							if (!MULTIVIEW_IMAGE_2_value.isEmpty()) {
								multiviewImagesMap.put(3, MULTIVIEW_IMAGE_2_value);
							}

							if (!MULTIVIEW_IMAGE_3_value.isEmpty()) {
								multiviewImagesMap.put(4, MULTIVIEW_IMAGE_3_value);
							}

							if (!MULTIVIEW_IMAGE_4_value.isEmpty()) {
								multiviewImagesMap.put(5, MULTIVIEW_IMAGE_4_value);
							}

							// Create Gift Card
							String giftCard = "FALSE";

							// Create Variant Image
							String variantImage = IMAGE_FULL_value;

							// Create Cost per item
							String costPerItem = "";
							Cell WHOLESALE_PRICE_US = sheet.getRow(i).getCell(columnNameMap.get("WHOLESALE_PRICE_US"));
							double WHOLESALE_PRICE_US_value = 0;
							if (WHOLESALE_PRICE_US != null
									&& WHOLESALE_PRICE_US.getCellType() == Cell.CELL_TYPE_NUMERIC) {
								WHOLESALE_PRICE_US_value = WHOLESALE_PRICE_US.getNumericCellValue();
								double costPerItemNumber = (WHOLESALE_PRICE_US_value * 1.3) + 25;
								costPerItem = "" + costPerItemNumber;
							} else {
								// System.out.println("WHOLESALE_PRICE_US is null in row "+i);
								isError = "true";
							}

							// Create Status
							String status = "active";

							// Create data0
							data0[0] = handle;
							data0[1] = title;
							data0[2] = html;
							data0[3] = vendor;
							data0[4] = type;
							data0[5] = tags;
							data0[6] = published;
							data0[7] = option1Name;
							data0[8] = option1Value;
							data0[9] = option2Name;
							data0[10] = option2Value[0];
							data0[13] = variantSKU;
							data0[14] = variantGrams;
							data0[15] = variantInventoryTracker;
							data0[16] = variantInventoryPolicy;
							data0[17] = variantFulfillmentService;
							data0[18] = variantPrice;
							data0[19] = variantCompareAtPrice;
							data0[20] = variantRequiresShipping;
							data0[21] = variantTaxable;
							data0[23] = IMAGE_FULL_value;
							data0[24] = "1";
							data0[26] = giftCard;
							data0[42] = variantImage;
							data0[45] = costPerItem;
							data0[46] = status;
							data0[47] = isError;
							datalist.add(data0);

							if (option2Value.length >= multiviewImagesMap.size()) {
								for (int l = 1; l < option2Value.length; l++) {
									// Create data1
									String[] data1 = new String[48];
									for (int m = 0; m < data1.length; m++) {
										data1[m] = "";
									}
									data1[0] = handle;
									data1[8] = option1Value;
									data1[10] = option2Value[l];
									data1[13] = variantSKU;
									data1[14] = variantGrams;
									data1[15] = variantInventoryTracker;
									data1[16] = variantInventoryPolicy;
									data1[17] = variantFulfillmentService;
									data1[18] = variantPrice;
									data1[19] = variantCompareAtPrice;
									data1[20] = variantRequiresShipping;
									data1[21] = variantTaxable;
									if (l < multiviewImagesMap.size() + 1) {
										data1[23] = multiviewImagesMap.get(l + 1);
										data1[24] = "" + (l + 1);
									}

									data1[42] = variantImage;
									data1[45] = costPerItem;
									datalist.add(data1);

								}
							} else if (option2Value.length < multiviewImagesMap.size()) {
								for (int l = 1; l < multiviewImagesMap.size(); l++) {
									// Create data1
									String[] data1 = new String[48];
									for (int m = 0; m < data1.length; m++) {
										data1[m] = "";
									}
									data1[0] = handle;
									data1[8] = option1Value;
									if (option2Value.length > l) {
										data1[24] = option2Value[l];
									}
									data1[13] = variantSKU;
									data1[14] = variantGrams;
									data1[15] = variantInventoryTracker;
									data1[16] = variantInventoryPolicy;
									data1[17] = variantFulfillmentService;
									data1[18] = variantPrice;
									data1[19] = variantCompareAtPrice;
									data1[20] = variantRequiresShipping;
									data1[21] = variantTaxable;
									data1[23] = multiviewImagesMap.get(l + 1);
									data1[24] = "" + (l + 1);
									data1[42] = variantImage;
									data1[45] = costPerItem;
									datalist.add(data1);
								}
							}
							if (isError.equalsIgnoreCase("True")) {
								System.out.println("NULL CELL in row " + (i + 1));
							}
						}
						if (i % 50 == 0) {
							Thread.sleep(100);
							appendExcelFile("products_shopify");
							System.out.println("Write======= " + (i + 1));
						}

					} catch (Exception e) {
						isError = "true";
						data0[47] = isError;
						datalist.add(data0);
						System.out.println("ERROR in row " + (i + 1));
						e.printStackTrace();
					}
				}

			}
			appendExcelFile("products_shopify");
			inputFile.close();
			System.out.println("$$$$$$FINSH$$$$$$");
		} catch (Exception ioe) {
			ioe.printStackTrace();
		}
	}

	public static double round(double value, int places) {
		if (places < 0)
			throw new IllegalArgumentException();

		long factor = (long) Math.pow(10, places);
		value = value * factor;
		long tmp = Math.round(value);
		return (double) tmp / factor;
	}

	private void appendExcelFile(String fileName) {
		Workbook workbook = null;
		Sheet sheet;
		try {
			File file = new File(path + fileName + ".xlsx");
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
			if (rowNum == 0) {
				String[] columns = { "Handle", "Title", "Body (HTML)", "Vendor", "Type", "Tags", "Published",
						"Option1 Name", "Option1 Value", "Option2 Name", "Option2 Value", "Option3 Name",
						"Option3 Value", "Variant SKU", "Variant Grams", "Variant Inventory Tracker",
						"Variant Inventory Policy", "Variant Fulfillment Service", "Variant Price",
						"Variant Compare At Price", "Variant Requires Shipping", "Variant Taxable", "Variant Barcode",
						"Image Src", "Image Position", "Image Alt Text", "Gift Card", "SEO Title", "SEO Description",
						"Google Shopping / Google Product Category", "Google Shopping / Gender",
						"Google Shopping / Age Group", "Google Shopping / MPN", "Google Shopping / AdWords Grouping",
						"Google Shopping / AdWords Labels", "Google Shopping / Condition",
						"Google Shopping / Custom Product", "Google Shopping / Custom Label 0",
						"Google Shopping / Custom Label 1", "Google Shopping / Custom Label 2",
						"Google Shopping / Custom Label 3", "Google Shopping / Custom Label 4", "Variant Image",
						"Variant Weight Unit", "Variant Tax Code", "Cost per item", "Status" };
				Row first = sheet.createRow(0);
				for (int i = 0; i < columns.length; i++) {
					first.createCell(i).setCellValue(columns[i]);
				}
				rowNum = rowNum + 1;
			}
			CellStyle errorStyle = workbook.createCellStyle();
			errorStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
			errorStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());

			for (String[] d : datalist) {
				Row row = sheet.createRow(rowNum++);

				for (int i = 0; i < d.length - 1; i++) {
					Cell cell = row.createCell(i);
					if (!d[d.length-1].isEmpty()) {
						cell.setCellStyle(errorStyle);
					}
					cell.setCellValue(d[i]);
				}
			}

			FileOutputStream fileOut = null;
			try {
				fileOut = new FileOutputStream(file);
				workbook.write(fileOut);
				datalist.clear();
			} catch (IOException e) {
				e.printStackTrace();
			} finally {
				if (fileOut != null) {
					try {
						fileOut.close();
					} catch (IOException e) {
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
		AndyEastoe panel = new AndyEastoe();
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