package excelConv;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;


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
	

	static final String tmp1 = "tmp1.csv";
	static final String tmp2 = "tmp2.csv";
	static final String xltmp = "input.xlsx";

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

//		Writer wr;
		try {
//			 br = new BufferedReader( new InputStreamReader( new FileInputStream(file),StandardCharsets.UTF_8));
			
			fr = new FileReader(file);
			br = new BufferedReader(fr);

//			wr = new FileWriter(new File(tmp2));
			FileOutputStream fileStream = new FileOutputStream(new File(tmp2));
			OutputStreamWriter wr = new OutputStreamWriter(fileStream, "Cp1255");;
//			FileWriter out = new FileWriter(new OutputStreamWriter(new FileOutputStream(outputXmlFilePath),"UTF-8"));

			
//			wr = new OutputStreamWriter(new FileOutputStream(tmp2), StandardCharsets.UTF_8);
			
			System.err.println("ENCODING"+fr.getEncoding());

			String str  = br.readLine();

			while(str != null){

				System.out.println(str);

				if(str.contains("³")){
					str = str.replaceAll("³", ",");
				}

				wr.write(str);
				wr.append('\n');

				str = br.readLine();
			}
			wr.flush();
			wr.close();
			fr.close();
			br.close();
			
			fileStream.close();


		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}


		try {
			DataConvertionUtil.csvToEXCEL(tmp2, xltmp);
		} catch (Exception e) {
			e.printStackTrace();
		}
		
		
		return xltmp;
	}

	
	public static void deleteTmps(){
		File f1 = new File(tmp1);
		File f2 = new File(tmp2);
		File f3 = new File(xltmp);
		
		try{
			if(f1.exists())
				f1.delete();
			if(f2.exists())
				f2.delete();
			if(f3.exists())
				f3.delete();
		}catch(Exception e){e.printStackTrace();}
	}


}
