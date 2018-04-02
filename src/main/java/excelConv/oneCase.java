package excelConv;

import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.Arrays;

import org.apache.poi.xssf.usermodel.XSSFSheet;

public class oneCase {

	String[][] unProcessedData;

	String Date="";
	String SampleID="";
	String Name="";
	String ID="";
	String SenderID="";
	String Testsname="";
	String Total="";
	String VAT=""; 
	String TotalWithVAT="";



	public oneCase(String[][] mat) {
		this.unProcessedData = mat;
	}

	public String[] toArray(){
		String[] str = {Date, SampleID, Name,ID,SenderID, Testsname, Total, VAT, TotalWithVAT};
		return str;
	}
	public void toSheet(XSSFSheet sheet){
		Funcs.StringArrToLastRow(toArray(), sheet, false);
	}

	public boolean process() {
		try{
			Date = unProcessedData[1][7].trim();
			SampleID = unProcessedData[1][6].trim();
			ID = unProcessedData[1][4].trim();
			SenderID = unProcessedData[1][3].trim();

			String name =  unProcessedData[1][5].trim().replaceAll("\\s+", " ");;
			if((name.toUpperCase().charAt(1) >= 65) && (name.toUpperCase().charAt(1) <= 90)){ //name in english
				this.Name = name;  
			}
			else{
				//				char[] carr = name.toCharArray();
				//				for(int i=carr.length-1; i>=0; i--){
				//					this.Name += carr[i]; 
				//				}
				this.Name = new StringBuffer(name).reverse().toString();

			}

			Testsname = unProcessedData[1][2].trim();
			for(int i=2; i<unProcessedData.length-2; i++){
				Testsname += " | "+ unProcessedData[i][2].trim();
			}

			TotalWithVAT = unProcessedData[unProcessedData.length-1][1].trim();

			// о"то + л"дс         47.60 о"то            280.00 л"дс                    
			String summ = unProcessedData[unProcessedData.length-1][2].trim().replaceAll("\\s+", " ");
			String[] arr = summ.split(" ");
			
			VAT = arr[3];
			Total = arr[5];

			try{
				Double d = Double.parseDouble(TotalWithVAT.replaceAll("\\s", ""));
//				System.out.println(d);
				d = Double.parseDouble(new DecimalFormat("#####.##").format(d));
//				System.out.println(d);
				TotalWithVAT = d+"0";
			}catch(Exception e){}

			System.out.println(Arrays.toString(toArray()));
		}catch(Exception e){System.err.println(e); return false;}


		return true;
	}

	public static void toFile(ArrayList<oneCase> proedcases, XSSFSheet mainSheet) {
		for(oneCase cs : proedcases){
			if(cs!=null){
				cs.toSheet(mainSheet);
			}
		}



	}
}
