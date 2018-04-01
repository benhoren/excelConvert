package excelConv;

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
		
	}
}
