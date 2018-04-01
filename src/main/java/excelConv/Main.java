package excelConv;


public class Main { 

	public static void main(String[] args) {

		System.out.println("start");
		String[][] sheet = Funcs.readExcel("explanation.xlsx");
		
		processData.play(sheet);
		
	}
}
