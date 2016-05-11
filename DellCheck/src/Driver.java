import java.io.IOException;
import java.util.ArrayList;

public class Driver {
	public static void main(String args[]) throws IOException{
		String excelFilepath = "CLS Poly +4 yr.xlsx";
		Computer computerTest = new Computer();
		
		ArrayList<Computer> testComputerlist = computerTest.readBooksFromExcelFile(excelFilepath);
		
		for(int i = 0; i < testComputerlist.size(); i++){
			System.out.println(i);
			System.out.println(testComputerlist.get(i).toString());
		}
		
	}
		
}
