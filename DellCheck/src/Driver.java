import java.io.IOException;
import java.util.ArrayList;

public class Driver {
	public static void main(String args[]) throws Exception{
		String excelFilepath = "CLS Poly +4 yr.xlsx";
		Computer computerTest = new Computer();
		
		ArrayList<Computer> testComputerlist = computerTest.readBooksFromExcelFile(excelFilepath);
		testComputerlist.remove(0);
		for(int i = 0; i < testComputerlist.size(); i++){
			System.out.println(i);
			testComputerlist.get(i).sendRequest();
			System.out.println(testComputerlist.get(i).toString());
		}
		
		testComputerlist.get(3).sendRequest();
		
	}
		
}
