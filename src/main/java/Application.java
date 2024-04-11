package main.java;

import java.awt.BorderLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTextArea;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.filechooser.FileSystemView;

import main.java.services.Converter;

public class Application extends JFrame implements ActionListener {

	/**
	 * @author parth ghavghave
	 */
	private static final long serialVersionUID = 1L;
	JButton button;
	JButton convertButton;
	JTextArea textArea;
	String inputFile = null;

	public Application() {
		setTitle("File Chooser Example");
		setSize(400, 200);
		setDefaultCloseOperation(EXIT_ON_CLOSE);

		button = new JButton("Choose File");
		button.addActionListener(this);

		convertButton = new JButton("Convert");
		convertButton.addActionListener(this);

		textArea = new JTextArea(10, 30);
		textArea.setEditable(false);

		JPanel panel = new JPanel();
		panel.add(button);
		panel.add(convertButton);

		JScrollPane scrollPane = new JScrollPane(textArea);

		getContentPane().add(panel, BorderLayout.NORTH);
		getContentPane().add(scrollPane, BorderLayout.CENTER);

		setLocationRelativeTo(null);
		setVisible(true);
	}

	public void actionPerformed(ActionEvent e) {
		if (e.getSource() == button) {
			JFileChooser fileChooser = new JFileChooser();
			FileSystemView fileSystemView = FileSystemView.getFileSystemView();
            File desktopDirectory = fileSystemView.getHomeDirectory();
            fileChooser.setCurrentDirectory(desktopDirectory);
            FileNameExtensionFilter filter = new FileNameExtensionFilter("Excel files (*.xls, *.xlsx)", "xls", "xlsx");
            fileChooser.setFileFilter(filter);
			int returnValue = fileChooser.showOpenDialog(this);
			if (returnValue == JFileChooser.APPROVE_OPTION) {
				File selectedFile = fileChooser.getSelectedFile();
				inputFile = selectedFile.getAbsolutePath();
				textArea.append("Selected file: " + selectedFile.getName() + "\n");
			}
		} else if (e.getSource() == convertButton) {
			if (inputFile == null) {
				textArea.append("File not selected\n");
			} else {
				textArea.append("Conversion started...\n");
				try {
					Converter.performConversion(inputFile);
					textArea.append("Conversion successfully completed :) \n **Output file stored on Desktop**\n\n");
				} catch (Exception e1) {
					textArea.append("Error occured...\nTry again with different file :)\n\n");
				}
			}
		}
	}

	public static void main(String[] args) {
		new Application();
	}
}
