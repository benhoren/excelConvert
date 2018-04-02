package excelConv;

import java.awt.EventQueue;

import javax.swing.JFrame;
import javax.swing.JFileChooser;
import javax.swing.SwingUtilities;
import javax.swing.UIManager;
import javax.swing.UnsupportedLookAndFeelException;
import javax.swing.filechooser.FileNameExtensionFilter;

import java.awt.BorderLayout;
import java.io.File;

import com.sun.java.swing.plaf.windows.WindowsFileChooserUI;

public class mainScreen {

	private JFrame frame;

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() { 
			public void run() {
				try {
					mainScreen window = new mainScreen();
					//					window.frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the application.
	 */
	public mainScreen() {
		initialize();
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		//		frame = new JFrame();
		//		frame.setBounds(100, 100, 450, 300);
		//		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		//		frame.getContentPane().setLayout(null);

		JFileChooser fileChooser = new JFileChooser();
		fileChooser.setCurrentDirectory(new File(System.getProperty("user.dir")));
//		FileNameExtensionFilter filter = new FileNameExtensionFilter("Excel File", "xlsx", "xlsm", "xls");
//		fileChooser.setFileFilter(filter);
		
		try {
			UIManager.setLookAndFeel("com.sun.java.swing.plaf.windows.WindowsLookAndFeel");
			SwingUtilities.updateComponentTreeUI(fileChooser);
			
	
		} catch (ClassNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (InstantiationException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IllegalAccessException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (UnsupportedLookAndFeelException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
//		this.pack();
		
		int result = fileChooser.showOpenDialog(frame);
		if (result == JFileChooser.APPROVE_OPTION) {
			File selectedFile = fileChooser.getSelectedFile();
			
			String path = selectedFile.getAbsolutePath();
			
			System.out.println("Selected file: " + path);
			
			Main.main(path);
			
			
//			String[][] sheet = Funcs.readExcel(selectedFile.getAbsolutePath());
//			processData.play(sheet);
			
			
		}


	}

}
