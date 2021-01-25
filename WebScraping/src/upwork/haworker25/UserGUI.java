package upwork.haworker25;

import javax.swing.*;
import javax.swing.filechooser.FileSystemView;
import javax.swing.text.TabableView;

import java.awt.event.*;
import java.awt.*;

public class UserGUI extends JPanel implements ActionListener {
	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;

	String url;
	JButton go;
	String folderPath="";
	static JTextArea ta;

	public UserGUI(String url) {
		JFrame f = new JFrame("WebScraping by haworker25");
		f.setSize(600, 400); 
		f.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE); 

		
		JPanel panel = new JPanel();
		JLabel label = new JLabel("Destination:");
		panel.add(label);
		JTextField folder = new JTextField(25);
		panel.add(folder);
		
		
		JButton chooseFolder = new JButton("Choose");
		chooseFolder.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent e) {
				JFileChooser chooser = new JFileChooser(FileSystemView.getFileSystemView().getHomeDirectory());
				chooser.setCurrentDirectory(new java.io.File("."));
				chooser.setDialogTitle("Choose destination");
				chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
				chooser.setAcceptAllFileFilterUsed(false);
				//
				if (chooser.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
					System.out.println("getCurrentDirectory(): " + chooser.getCurrentDirectory());
					folderPath = chooser.getSelectedFile().getAbsolutePath();
					folder.setText(folderPath);
				} else {
					System.out.println("No Selection ");
				}
				
			}
		});
		panel.add(chooseFolder);	
		go = new JButton("Extract");
		go.addActionListener(this);
		panel.add(go);
		
		
		ta = new JTextArea();
		
		f.getContentPane().add(BorderLayout.NORTH,panel);
		f.getContentPane().add(BorderLayout.CENTER,ta);
		f.setVisible(true);
	}

	public void actionPerformed(ActionEvent e) {
		if(folderPath.isEmpty()) {
			JOptionPane.showMessageDialog(null, "Please choose destination folder!");
			return;
		}else {
			UserGUI.setLog("========Start===========");
			Scraping scraping = new Scraping(url,folderPath);
			scraping.processWeb();
			
		}
	}

	public static void setLog(String log) {
		String currentLog= ta.getText();
		currentLog=currentLog+"\n"+log;
		ta.setText(currentLog);
	}
}