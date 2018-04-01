package excelConv;

import java.util.Arrays;

public class Main { 

	public static void main(String[] args) {
		// TODO Auto-generated method stub

		System.out.println("start");
		String[][] sheet = Funcs.readMatrix("explanation.xlsx");
		
		processData.play(sheet);
		

	}

}
