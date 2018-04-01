package excelConv;

import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Funcs {


	/**
	 * function to process a full excel sheet to string matrix
	 * @param path path to excel file
	 * @return string matrix of the first sheet.
	 */
	public static String[][] readExcel(String path){
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
						//asheet.getRow(i).getCell(j).setCellType(Cell.CELL_TYPE_STRING);
						//sheet[i][j] = asheet.getRow(i).getCell(j).getStringCellValue();
						Cell cell = asheet.getRow(i).getCell(j);
						if (cell != null) {
							String strCellValue = "";

							switch (cell.getCellType()) {
							case Cell.CELL_TYPE_STRING:
								strCellValue = cell.toString();
								break;
							case Cell.CELL_TYPE_NUMERIC:
								if (DateUtil.isCellDateFormatted(cell)) {
									SimpleDateFormat dateFormat = new SimpleDateFormat(
											"dd/MM/yyyy");
									strCellValue = dateFormat.format(cell.getDateCellValue());
								} else {
									cell.setCellType(Cell.CELL_TYPE_STRING);

									strCellValue = cell.getStringCellValue();
								}
								break;
							case Cell.CELL_TYPE_BOOLEAN:
								strCellValue = new String(new Boolean(
										cell.getBooleanCellValue()).toString());
								break;
							case Cell.CELL_TYPE_BLANK:
								strCellValue = "";
								break;
							}
							sheet[i][j] = strCellValue;
						}




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

		//Date	Sample ID	Name	ID	Sender ID	Tests name	Total	VAT	Total With VAT

		String taz = "ת"+'"'+"ז";
		String sach = "סה"+'"'+"כ";
		String maam = "מע"+'"'+"מ";
		String[] ques={"תאריך","מס' בדיקה","שם",taz,"מס' תלוש","בדיקה",sach, maam ,sach+ " כולל "+ maam};
		Funcs.StringArrToLastRow(ques, mainSheet, true);

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

	public static void openFile(String filename){
		try {
			Desktop.getDesktop().edit(new File(filename+".xlsx"));
		} catch (IOException e) {
			// TODO Auto-generated catch block
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
