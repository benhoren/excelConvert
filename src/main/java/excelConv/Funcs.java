package excelConv;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Funcs {


	static final String outputFold = "output";

	/**
	 * function to process a full excel sheet to string matrix
	 * @param path path to excel file
	 * @return string matrix of the first sheet.
	 */
	public static String[][] readMatrix(String path){
		String[][] sheet = null;
		FileInputStream file;
		try {
			file = new FileInputStream(new File(path));
			XSSFWorkbook workbook = new XSSFWorkbook(file);
			XSSFSheet asheet = workbook.getSheetAt(0);
			System.out.println("LR: "+asheet.getLastRowNum());
			sheet = new String[asheet.getLastRowNum()+1][];
			
			for(int i=0; i<=asheet.getLastRowNum(); i++){

				if(asheet.getRow(i) == null){
					sheet[i] = null;
					continue;
				}

				sheet[i] = new String[asheet.getRow(i).getLastCellNum()+1];
				
				for(int j=0; j<=asheet.getRow(i).getLastCellNum(); j++){

					if(asheet.getRow(i).getCell(j)==null){
						sheet[i][j] = "";
						continue;
					}
					try{
						sheet[i][j] = asheet.getRow(i).getCell(j).getStringCellValue();
					}catch(Exception e){System.err.println(e);
					sheet[i][j] = ""+asheet.getRow(i).getCell(j).getNumericCellValue();
					}
				}

			}

		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		return sheet;
	}
	
	
	
	static XSSFWorkbook workbook;
	static FileOutputStream outputStream;
	static XSSFSheet mainSheet;

	public static void startWriters(){
		workbook = new XSSFWorkbook();

		mainSheet = workbook.createSheet("summary");
		
//Date	Sample ID	Name	ID	Sender ID	Tests name	Total	VAT	Total With VAT

		String taz = "ת"+'"'+"ז";
		String sach = "סה"+'"'+"כ";
		String maam = "מע"+'"'+"מ";
		String[] ques={"תאריך","מס' בדיקה","שם",taz,taz +" שולח","שמות בדיקות",sach, maam ,sach+ " כולל "+ maam};
		Funcs.StringArrToLastRow(ques, mainSheet);

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
	public static void StringArrToLastRow(String[] arr, XSSFSheet sheet) {
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

		int i=0;
		Cell cell;
		for(i=0; i<arr.length; i++){
			if(arr[i].length()>32000){
				System.err.println("EXPAND ERROR");
				arr[i] = arr[i].substring(0, 32000);
			}
			cell = row.createCell(i);
			try {
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
