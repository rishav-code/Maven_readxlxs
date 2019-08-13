package Atmecs.xlsread;

import java.io.File;
import java.io.FileInputStream;
import java.sql.SQLException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class read {
	public static void main(String[]args) throws SQLException{
	
	try {
		String fle="C:\\Users\\rishav.kumar\\Desktop\\details.xlsx";
		File fi=new File(fle);
	    
	    XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(fi));
	    XSSFSheet sheet = wb.getSheetAt(0);
	    XSSFRow row;
	    XSSFCell cell;

	    int rows; // No of rows
	    rows = sheet.getPhysicalNumberOfRows();

	    int cols = 0; // No of columns
	    int tmp = 0;

	    
	    for(int i = 0;  i < rows; i++) {
	        row = sheet.getRow(i);
	        if(row != null) {
	            tmp = sheet.getRow(i).getPhysicalNumberOfCells();
	            if(tmp > cols) cols = tmp;
	        }
	    }

	    for(int r = 0; r < rows; r++) {
	        row = sheet.getRow(r);
	        if(row != null) {
	            for(int c = 0; c < cols; c++) {
	                cell = row.getCell((short)c);
	                if(cell != null) {
	                	
	                
	                   
	                	System.out.print(cell+" ");
	                }
	            }
	            System.out.println();
	        }
	    }
	} catch(Exception ioe) {
	    ioe.printStackTrace();
	}

}
}

