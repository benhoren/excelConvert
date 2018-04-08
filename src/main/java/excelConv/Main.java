package excelConv;

import java.awt.Desktop;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Scanner;
import java.util.stream.Stream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class Main {

	//	public static void main(String[] args) {
	//
	//		String filename = "explanation.xlsx";
	//		main(filename);
	//		
	//	}

	public static void main(String filename){
		System.out.println("start");
		String inputxl = unknown_Csv_Excel(filename);

		String[][] sheet = Funcs.readExcel(inputxl);
		processData.play(sheet);

				deleteTmps();
	}


	static final String tmp1 = "tmp1";
	static String xltmp = "input";

	public static void copyFile( File from, File to ) throws IOException {
		Files.copy( from.toPath(), to.toPath() );
	}


	/**
	 * from unknown file type, to .csv, expand the csv and then from csv to an excel file .xlsx
	 * @param path
	 */
	public static String unknown_Csv_Excel(String path){
		File file = new File(path);

		String oldname = path;
		file.renameTo(new File(tmp1));
		file = new File(tmp1);

		File dirFrom = new File(tmp1);
		File dirTo = new File(oldname);
		try {
			copyFile(dirFrom, dirTo);
		} catch (IOException ex) {
			ex.printStackTrace(); 
		}




		FileReader fr;
		BufferedReader br;


		try {


//			FileInputStream fis = new FileInputStream(file);
//			InputStreamReader isr = new InputStreamReader(fis, Charset.forName("ISO-8859-1"));
//			System.out.println(isr.getEncoding());
//			try
//			{
//				while (isr.ready())
//				{
//					System.out.print("" + (char) isr.read());
//				}
//			} catch (IOException e)
//			{
//				throw e;
//			} finally
//			{
//
//				isr.close();
//			}

			//			fr = new FileReader(file);
			//			br = new BufferedReader(fr);
			//			br  = new BufferedReader(new InputStreamReader(new FileInputStream(file), "Cp1252"));
			//			 br = new BufferedReader(new InputStreamReader(new FileInputStream(file), Charset.forName("ISO-8859-1")));
			//			Scanner in = new Scanner(new FileReader(file));
						br = new BufferedReader (new InputStreamReader(new FileInputStream(file),Charset.forName("iso-8859-8")));
						xltmp = readyFiles(xltmp);
//						System.out.println(br.as);
						
					
//						OutputStream youInputStream ;
//						Writer out = new OutputStreamWriter(youInputStream , "UTF-8");
//
//						out.write(yourText);


						

			//			System.err.println("ENCODING"+fr.getEncoding());

						String str  = br.readLine();
			
						while(str != null){
			
							System.out.println(str);
			
//							String decodedToUTF8 = new String(str.getBytes("ISO-8859-1"), "UTF-8");
	
//							System.out.println(decodedToUTF8);
							
							if(str.contains("³")){
								str = str.replaceAll("³", ",");
							}
							StringArrToLastRow(str.split(","),mainSheet, false);
							
							
							str = br.readLine();
						}
				
			//			fr.close();
						br.close();
						
						closeWriters();
			
			
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		//
		//
		////		try {
		////			DataConvertionUtil.csvToEXCEL(tmp2, xltmp);
		////		} catch (Exception e) {
		////			e.printStackTrace();
		////		}
		//		
				xltmp = xltmp+".xlsx";
		return xltmp;
	}


	public static void deleteTmps(){
		File f1 = new File(tmp1);
		File f3 = new File(xltmp);

		try{
			if(f1.exists())
				f1.delete();
			if(f3.exists())
				f3.delete();
		}catch(Exception e){e.printStackTrace();}
	}




	static XSSFWorkbook workbook;
	static FileOutputStream outputStream;
	static XSSFSheet mainSheet;

	public static String readyFiles(String fileName){
		boolean ok =false;
		int i=1;

		startWriters();

		File directory = new File(fileName+".xlsx");
		if (!directory.exists()){
			try {
				outputStream = new FileOutputStream(fileName+".xlsx");
				return fileName;
			} catch (FileNotFoundException e) {}
		}

		else{
			ok = false;
			int j=1;
			while(!ok){
				directory = new File(fileName+"-"+j+".xlsx");
				if (!directory.exists()){
					try {
						outputStream = new FileOutputStream(fileName+"-"+j+".xlsx");
						ok =true;
						return fileName+"-"+j;
					} catch (FileNotFoundException e) {}
				}
				j++;
			}
		}
		return "";
	}

	public static void startWriters(){
		workbook = new XSSFWorkbook();
		mainSheet = workbook.createSheet("summary");
	}



	public static void closeWriters(){
		try {
			workbook.write(outputStream);
			outputStream.close();
			workbook.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}



	/**
	 * this function write a String arr to the last row in sheet
	 * @param arr
	 * @param sheet
	 */
	public static void StringArrToLastRow(String[] arr, XSSFSheet sheet, boolean bold) {
		if(arr == null) return;

		for(int i=0; i<arr.length; i++){
			if(arr[i] == null)
				arr[i]="";
		}

		Row row = sheet.getRow(sheet.getLastRowNum());
		if(row==null)
			row = sheet.createRow(sheet.getLastRowNum());
		else row = sheet.createRow(sheet.getLastRowNum()+1);

		for(int i=0 ;i<arr.length; i++){
			if(arr[i].length()>30000){
				expand(arr);
			}
		}

		for(int i=0 ;i<arr.length; i++){
			if(arr[i].length()>30000){
				expand(arr);
			}
		}

		XSSFCellStyle style = (XSSFCellStyle) workbook.createCellStyle();
		Font font = workbook.createFont();
		font.setFontHeight((short)(12*20));
		font.setFontName("Arial");
		font.setBold(true);
		style.setFont(font);

		int i=0;
		Cell cell;
		for(i=0; i<arr.length; i++){
			if(arr[i].length()>32000){
				System.err.println("EXPAND ERROR");
				arr[i] = arr[i].substring(0, 32000);
			}
			cell = row.createCell(i);
			try {
				if(bold)
					cell.setCellStyle(style);
				cell.setCellValue(arr[i]);
			}catch(Exception e) {e.printStackTrace();}
		}
	}

	/**
	 * expand big arr cells to next cells
	 * @param arr array with Strings.
	 */
	private static void expand(String[] arr){
		String tmp = "";
		for(int i=0; i<arr.length-1; i++){
			if(arr[i].length()>30000){
				tmp = arr[i].substring(30000, arr[i].length());
				arr[i] = arr[i].substring(0, 30000);
				arr[i+1] = tmp;
			}
			tmp="";
		}
	}
}
