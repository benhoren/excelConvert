package excelConv;

import java.util.ArrayList;


public class processData {

	static String[][] summary; 
	
	public static void play(String[][] sheet){
		fixData(sheet);
		printSheet(sheet);
		ArrayList <oneCase> unprocases= splitToCases(sheet);
	}


	private static ArrayList<oneCase> splitToCases(String[][] sheet) {

		//find headers row
		int start = -1;
		for(int i=0; i<sheet.length; i++){
			if(sheet[i]==null)
				continue;
			for(int j=0; j<sheet[i].length; j++){
				if(sheet[i][j].trim().equals("ךיראת")){
					System.out.println("first row at "+i);
					start = i;
					break;
				}
			}
			if(start!=-1) break;
		}



		ArrayList<oneCase> cases= new ArrayList<oneCase>();

		//split to cases
		String[][] mat = null;
		oneCase oc = null;
		int count = 1;
		mat = null;
		boolean ok = true;
		for(int i=start+2; i<sheet.length; i++){
			
			if((sheet[i]!=null)&&(sheet[i][0]!=null)&&(sheet[i][0].contains("****")))
				break;

			ok = false;

			if(sheet[i] == null) ok = true;
			else{
				String str = sheet[i][1];
				if(str.contains("-----")){
					i+=2;
					count+=2;
					ok = true;
				}
			}

			if(ok){
				if(count>1){
					mat = new String[count][];
					for(int k=0; k<count; k++){
						mat[k] = sheet[i-count+k];
					}
					count = 1;
					oc = new oneCase(mat);
					cases.add(oc);

					oc = null;
					mat= null;
				}
			}
			else {
				count++;
				boolean flag = false;
				for(int k=0; k<sheet[i].length; k++){
					if(sheet[i][k].trim().equals("ךיראת")){
						System.out.println("head again");
						i++;
						count = 1;
						break;
						
					}
				}
			}
		}

		summary = cases.get(cases.size()-1).unProcessedData;
		cases.remove(cases.size()-1);
		
		System.out.println();
		System.out.println("size: "+cases.size());
		System.out.println();

		for(oneCase cs : cases){
			System.out.println("****************");
			mat = cs.unProcessedData;
			
			if(mat == null) {System.out.println("null");continue;}
			printSheet(mat);

		}

		

		return cases;

	}


	public static void fixData(String[][] sheet){

		for(int i=0; i<sheet.length; i++){
			boolean ok = true;

			if(sheet[i] == null)
				continue;

			for(int j=0; j<sheet[i].length; j++){
				if(sheet[i][j]==null)
					continue;
				if(sheet[i][j].trim().isEmpty())
					continue;

				ok = false;
				break;
			}

			if(ok){
				sheet[i] = null;
			}

		}

	}


	public static void printSheet(String[][] sheet){

		System.out.println(sheet.length);
		for(int i=0; i<sheet.length; i++){
			System.out.println();
			System.out.print("row "+i+": ");
			if(sheet[i] == null) {System.out.println("null");continue;}
			for(int j=0; j<sheet[i].length; j++){
				System.out.print(sheet[i][j]+", ");
			}
		}
		System.out.println();
		System.out.println("end");
	}

}
