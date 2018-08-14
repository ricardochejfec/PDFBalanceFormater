import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class MBDocMaker {
	
	private String sourceFileName; 
	private String destDirectory;
	private String month;
	private List<String> ids = new ArrayList<String>();
	private List<String> balances = new ArrayList<String>();
	private List<Integer> specialRows = new ArrayList<Integer>();
	private int idsColumn; 
	private int balancesColumn;
	private int dataLength;
	private String idsName = "ID Number";
	private String balancesName = "Latest Balance FY19";
//	private String idsName = "I.D.";
//	private String balancesName = "Balance";
//  somehow not writting the first homie
	
	public MBDocMaker(String sourceFileName, String destDirectory, String month) throws EncryptedDocumentException, InvalidFormatException, IOException {
		this.month = month;
		this.sourceFileName = sourceFileName; 
		this.destDirectory = destDirectory;
		this.isolateColumns();
		this.dataLength = ids.size();
	}
	
	
	@SuppressWarnings("deprecation")
	public void isolateColumns() throws EncryptedDocumentException, InvalidFormatException, IOException {
		
		Workbook workbookIn = null;

		workbookIn = WorkbookFactory.create(new File(this.sourceFileName));
		     // Getting the Sheet at index zero
		Sheet sheet = workbookIn.getSheetAt(0);
		
		 // Create a DataFormatter to format and get each cell's value as String
		DataFormatter dataFormatter = new DataFormatter(); 
		
		int colIndex;
		int rowIndex = 1; 
		int a = 0; 
		int b = 0;
		
//		System.out.println("\nIterating over Rows and Columns using for-each loop\n");
        for (Row row: sheet) {
        	colIndex=1;
	         for(Cell cell: row) {
	         	String cellValue = dataFormatter.formatCellValue(cell);
	         	if (rowIndex <=2) {
//	         		System.out.println(cellValue);
	         		findColumnNumbers(cellValue, colIndex);
	         		colIndex++;
	         	}
	         	else  {     	
	         		if (colIndex==idsColumn){
//	         			System.out.println("A- "+a);
	         			a++;
	         			this.ids.add(cellValue);
	         		} else if (colIndex==balancesColumn){
//	         			System.out.println("B- "+b);
	         			b ++;
	         			if(cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
//	         		        System.out.println("Formula is " + cell.getCellFormula());
	         		        switch(cell.getCachedFormulaResultType()) {
	         		            case Cell.CELL_TYPE_NUMERIC:
//	         		                System.out.println("Last evaluated as: " + cell.getNumericCellValue());
	         		            	this.balances.add("$" + Math.round(cell.getNumericCellValue() * 100.0) / 100.0); //fix for decimals 
	         		                break;
	         		            case Cell.CELL_TYPE_STRING:
//	         		                System.out.println("Last evaluated as \"" + cell.getRichStringCellValue() + "\"");
	         		            	this.balances.add("$"+cell.getRichStringCellValue()); 
	         		                break;
	         		        }
	         		     }
	         			else {
	         				this.balances.add(cellValue);
	         			}
	         			
	         		}
	         		colIndex++;
	         	}
	         }rowIndex++;
        }
//        System.out.println(""+ idsColumn +" " + balancesColumn);
//        System.out.println(balances.toString());
//        System.out.println("" + ids.size() +  " " + balances.size());
	}
	
	public String createDoc(){
		int repeatFactor = 30;
		int height = (dataLength/3) + 30; 
	
		
		String fileName = month + "- Monthly Balances.xlsx";
		
		List<ArrayList<String>> reorderedData = reorderColumns(repeatFactor);
	    
		//Create file system using specific name
	    FileOutputStream out = null;
		try {
			String fullFname =  destDirectory + "/" + fileName;
			out = new FileOutputStream(new File(fullFname));
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			return (e.getMessage());
		}
		
		XSSFWorkbook workbookOut = new XSSFWorkbook();
		XSSFSheet spreadsheet = workbookOut.createSheet(fileName);
		
		//Create row object
	    XSSFRow row;
	    Map<String, Object[]> empinfo = new  LinkedHashMap<String, Object[]>();
	    
	    //	    i < (reorderedData.get(0).size()+(length/repeatFactor))
	    for(int i = 1; i < height ; i+=(repeatFactor+1)){
//	    	System.out.println(""+i);
		    empinfo.put( "" + i, new Object[] {"", "I.D.", "Balance", "", "I.D.", "Balance", "", "I.D.", "Balance" });
		    for(int j = i+1; j < (i+repeatFactor+1); j++){
//		    	System.out.println("J-"+j);
		    	if (reorderedData.get(0).size() == 0){
		    		String rowNum = "" + j;
			    	empinfo.put(rowNum, new Object[] {"", "", "", "", "", "", "", ""});
		    	}
		    	else {
		    		String rowNum = "" + j;
		    		empinfo.put(rowNum, new Object[] {"", reorderedData.get(0).remove(0), reorderedData.get(1).remove(0), "", reorderedData.get(2).remove(0), reorderedData.get(3).remove(0), "", reorderedData.get(4).remove(0), reorderedData.get(5).remove(0) });
		    	}
		    }
	    }
	    
	    
	    CellStyle regStyle = workbookOut.createCellStyle();
	    regStyle.setAlignment(HorizontalAlignment.CENTER);
	    
	    CellStyle boldedStyle = workbookOut.createCellStyle();
	    Font boldedFont = workbookOut.createFont();
	    
	    boldedFont.setBold(true);
	    boldedFont.setFontHeight((short) (boldedFont.getFontHeight()+4));

	    boldedStyle.setFont(boldedFont);
	    boldedStyle.setAlignment(HorizontalAlignment.CENTER);
	    boldedStyle.setVerticalAlignment(VerticalAlignment.CENTER);
 
	    Set<String> keyid = empinfo.keySet();
	
	    int rowid = 0;

	    for (int i = 1; i < keyid.size(); i+=31){
	    	specialRows.add(i);
	    }
	    
	    for (String key : keyid){
	    	row = spreadsheet.createRow(rowid++);
	    	System.out.println(rowid);
	        Object [] objectArr = empinfo.get(key);
	        int cellid = 0;
	        for (Object obj : objectArr){
	        	Cell cell = row.createCell(cellid++);
	        	cell.setCellValue((String)obj);
	        	if(specialRows.contains(rowid)){
	        		row.setHeightInPoints((2*spreadsheet.getDefaultRowHeightInPoints()));
	        		cell.setCellStyle(boldedStyle);
	        	}
	        	else {
	        		cell.setCellStyle(regStyle);
	        	}
	        }
	    }
	    
	    spreadsheet.autoSizeColumn(1); 
	    spreadsheet.autoSizeColumn(2);
	    spreadsheet.autoSizeColumn(4); 
	    spreadsheet.autoSizeColumn(5); 
	    spreadsheet.autoSizeColumn(7); 
	    spreadsheet.autoSizeColumn(8); 
	    
	    //write operation workbook using file out object 
	    try {
	    	workbookOut.write(out);
			out.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			return (e.getMessage());
		}
	    return "Document Created Succcesfully";
	}

	public void findColumnNumbers(String cellValue, int index){
		if(cellValue.equals(idsName)){
 			this.idsColumn = index;
 		}
 		else if(cellValue.equals(balancesName)){
 			this.balancesColumn = index; 
 		}
	}
	
	
	public List<ArrayList<String>> reorderColumns(int repeatFactor){
	
		int cellsPerSection = repeatFactor*3;
		List<ArrayList<String>> reorderedData = new ArrayList<ArrayList<String>>();
		
		for (int i = 0; i < 6; i++){
			reorderedData.add(new ArrayList<String>());
		}
		
		for (int i = 0; i < dataLength; i++){
			if (!ids.isEmpty()){
				if ((i % cellsPerSection) < repeatFactor){
					reorderedData.get(0).add(ids.remove(0));
					reorderedData.get(1).add(balances.remove(0));
				} else if (((i % cellsPerSection) >= repeatFactor) && ((i % cellsPerSection) < repeatFactor*2)){
					reorderedData.get(2).add(ids.remove(0));
					reorderedData.get(3).add(balances.remove(0));
				} else if ((i % cellsPerSection) >= repeatFactor*2){
					reorderedData.get(4).add(ids.remove(0));
					reorderedData.get(5).add(balances.remove(0));
				} 
			}
		}
		
		for(int i = 2; i<6;i++){
			for(int j = 0; j<reorderedData.get(0).size();j++){
				try {
					reorderedData.get(i).get(j);
				}
				catch (IndexOutOfBoundsException e){
					reorderedData.get(i).add("");
				}
			}
		}
		
		//System.out.println(reorderedData.toString());
		return reorderedData;
	}
	
	public String toString(){
		return ("IDS ->" + this.ids.toString() + "\nBalances ->" + this.balances.toString());
		
	}

}
