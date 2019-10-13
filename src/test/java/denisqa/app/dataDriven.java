package denisqa.app;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class dataDriven {

	//Identify testcases column by scanning he entire row
	//Once column is identified then scan entire testcases column to grab "Purchase" testcase
	//After u grab "Purchase" testcases row- pull all the data of that row and feed into test
	
	//public ArrayList<String> getData(String testcaseName) throws IOException {
		
	//fileInputStream argument
	public ArrayList<String> getData(String testcaseName) throws IOException {
		
		ArrayList<String> ar = new ArrayList<String>();
		
		FileInputStream fis = new FileInputStream("C://Users//dmarkov//Desktop//SomeRandomExcel.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		
		int sheets = workbook.getNumberOfSheets();
		for (int i = 0; i < sheets; i++) {
			if (workbook.getSheetName(i).equalsIgnoreCase("testdata")) {
				
			XSSFSheet sheet = workbook.getSheetAt(i);
			//Identify testcases column by scanning he entire row
			Iterator<Row> rows = sheet.iterator(); //sheet is collection of rows
			Row firstrow = rows.next();
			Iterator<Cell> ce = firstrow.cellIterator(); //row is collection of cells
			int k = 0;
			int column = 0;
			
		while(ce.hasNext()) {
			Cell value = ce.next();
				if (value.getStringCellValue().equalsIgnoreCase("Testcases")) {
					//desired column
					column = k;
				}
				k++;
		}
			System.out.println(column);
			
		while(rows.hasNext()) {
			Row r = rows.next();
			if (r.getCell(column).getStringCellValue().equalsIgnoreCase(testcaseName)) {
				//pull all the data of that row and feed into test
				
				Iterator<Cell> cv = r.cellIterator();
				while(cv.hasNext()) {
					Cell c = cv.next();
					if (c.getCellType()==CellType.STRING) {
						ar.add(c.getStringCellValue());
					}
					else {
						ar.add(NumberToTextConverter.toText(c.getNumericCellValue()));
						
					}
				}
			}
			
			
		}
			
			}
		}
		return ar;
	}
	
	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		
		
		
}
		
	}


