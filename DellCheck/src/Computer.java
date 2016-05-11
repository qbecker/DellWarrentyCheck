
import java.awt.List;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Computer {
	private String propertyControlNumber;
	private String owner;
	private String makeModel;
	private String serviceTag;
	
	public Computer(){}
	
	
	
	public String getPropertyControlNumber() {
		return propertyControlNumber;
	}



	public void setPropertyControlNumber(String propertyControlNumber) {
		this.propertyControlNumber = propertyControlNumber;
	}



	public String getOwner() {
		return owner;
	}



	public void setOwner(String owner) {
		this.owner = owner;
	}



	public String getMakeModel() {
		return makeModel;
	}



	public void setMakeModel(String makeModel) {
		this.makeModel = makeModel;
	}



	public String getServiceTag() {
		return serviceTag;
	}



	public void setServiceTag(String serviceTag) {
		this.serviceTag = serviceTag;
	}



	public Object getCellValue(Cell cell){
		switch(cell.getCellType()){
		case Cell.CELL_TYPE_STRING:
			return cell.getStringCellValue();
		
		case Cell.CELL_TYPE_BOOLEAN:
			return cell.getBooleanCellValue();
			
		case Cell.CELL_TYPE_NUMERIC:
			return cell.getNumericCellValue();
		}
		return null;
	}

	public ArrayList<Computer> readBooksFromExcelFile(String excelFilePath) throws IOException {
		ArrayList<Computer> computerList = new ArrayList<>();
		 FileInputStream inputStream = new FileInputStream(new File(excelFilePath));   
	     Workbook workbook = new XSSFWorkbook(inputStream);
	     Sheet firstSheet = workbook.getSheetAt(0);
	     Iterator<Row> iterator = firstSheet.iterator();
	     
	     
		  while (iterator.hasNext()){
			Row nextRow = iterator.next();
			Iterator<Cell> cellIterator = nextRow.cellIterator();
			Computer aComputer = new Computer();
			
			while(cellIterator.hasNext()){
				Cell nextCell = cellIterator.next();
				int columIndex = nextCell.getColumnIndex();
				
				switch(columIndex){
				case 0:
					aComputer.setPropertyControlNumber(nextCell.toString());
					break;
				case 1:
					aComputer.setOwner(nextCell.toString());
					break;
				case 2: 
					aComputer.setMakeModel(nextCell.toString());
					break;
				case 3:
					aComputer.setServiceTag(nextCell.toString());
					break;
				}
			}
			computerList.add(aComputer);
		}
		 workbook.close();
		 inputStream.close();
		return computerList;
	}
	
	public String toString(){
		String result = "";
		result += "Owner: " + getOwner() + "\n";
		result += "Make/Model: " + getMakeModel() + "\n";
		result += "Property Control Number: " + getPropertyControlNumber() + "\n";
		result += "Service Tag:  " + getServiceTag() + "\n";
		return result;
	}

}
