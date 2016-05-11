
import java.awt.List;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.MalformedURLException;
import java.net.URL;
import java.net.URLConnection;
import java.util.ArrayList;
import java.util.Iterator;

import java.io.StringWriter;


import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;

import org.w3c.dom.Document;
import org.w3c.dom.NodeList;

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
	private String shippingDate;
	
	public Computer(){}
	
	public String getShippingDate(){
		return shippingDate;
	}
	
	public void setShippingDate(){
		this.shippingDate = shippingDate;
	}
	
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
	
	public void sendRequest() throws Exception{
		
		URL url = new URL("https://api.dell.com/support/v2/assetinfo/warranty/tags.xml?svctags=4ZGZLG1&apikey=849e027f476027a394edd656eaef4842");
		URLConnection connection = url.openConnection();
		
		Document doc = parseXML(connection.getInputStream());
		NodeList descNodes = doc.getChildNodes();
		for(int i = 0; i<descNodes.getLength(); i++){
			System.out.println(descNodes);
		}
		
		
		
		
		
		
		System.out.println(url);
		
	}
	
	private Document parseXML(InputStream stream) throws Exception {
		DocumentBuilderFactory objDocumentBuilderFactory = null;
        DocumentBuilder objDocumentBuilder = null;
        Document doc = null;
        try
        {
            objDocumentBuilderFactory = DocumentBuilderFactory.newInstance();
            objDocumentBuilder = objDocumentBuilderFactory.newDocumentBuilder();

            doc = objDocumentBuilder.parse(stream);
        }
        catch(Exception ex)
        {
            throw ex;
        }       

        return doc;
    
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
