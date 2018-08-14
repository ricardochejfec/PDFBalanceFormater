import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Random;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class MBTesterSSCreator {
	
	public static void main(String[] args) throws IOException {
		
		Random rand = new Random();
		//Create blank workbook
	    XSSFWorkbook workbook = new XSSFWorkbook(); 
	    //Create a blank sheet
		XSSFSheet spreadsheet = workbook.createSheet("Clients");
		//Create row object
	    XSSFRow row;
	    Map < String, Object[] > empinfo = new TreeMap < String, Object[] >();
	    empinfo.put( "1", new Object[] { "Name", "Last Name", "I.D.", "Balance", "CellPhone" });
	    for(int i = 2; i < 302; i++){
	    	empinfo.put( "" + i, new Object[] { "" + (i-1), "Bob", "" + (i-1), ""+rand.nextInt(10000),  ""+rand.nextInt(100000) });
	    	System.out.println(""+(i-1));
	    }
	    Set < String > keyid = empinfo.keySet();
	      int rowid = 0;
	      for (String key : keyid)
	      {
	         row = spreadsheet.createRow(rowid++);
	         Object [] objectArr = empinfo.get(key);
	         int cellid = 0;
	         for (Object obj : objectArr)
	         {
	            Cell cell = row.createCell(cellid++);
	            cell.setCellValue((String)obj);
	         }
	      }
	      //Write the workbook in file system
	      FileOutputStream out = new FileOutputStream( 
	      new File("Writesheet.xlsx"));
	      workbook.write(out);
	      out.close();
	      System.out.println( 
	      "Writesheet.xlsx written successfully" );
	   }
}




