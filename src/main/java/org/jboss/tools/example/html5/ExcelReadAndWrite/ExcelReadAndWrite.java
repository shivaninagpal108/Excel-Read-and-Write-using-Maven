package org.jboss.tools.example.html5.ExcelReadAndWrite;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.Iterator;
import java.util.Map;
import java.util.HashMap;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class ExcelReadAndWrite 
{
	public static void main( String[] args )
    {
    	try
    	{
    		 File excel = new File("MyExcel.xlsx");
    	     if (!excel.exists()) 
    	     {
    	    	 System.out.println("File does not exist");
    	         if (!excel.createNewFile())
    	        	 System.out.println("File cannot be created");
    	         else
    	             System.out.println("File created");
    	     } 
    	     else 
    	     {
    	    	 System.out.println("File exists");
    	    	 if(!excel.canRead())
    	    		 System.out.println("Error in reading. Need permission");
    	         if(!excel.canWrite())
    	        	 System.out.println("Error in writing. Need permission");
    	     }
             XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(excel));
             System.out.println("Workbook has " + workbook.getNumberOfSheets() + " Sheets.");
             //Reading from Excel and printing to console and writing to an Map
             XSSFSheet sheetRead = workbook.getSheetAt(0);
             Map<Integer, Object[]> excel_data = new HashMap<Integer, Object[]>();
             int key = 1;
             Iterator<Row> rowItr = sheetRead.iterator();
             while (rowItr.hasNext()) 
             {
                 Row row = rowItr.next();
                 Object[] cellData = new Object[row.getLastCellNum()];
                 int i = 0;
                 Iterator<Cell> cellIterator = row.cellIterator();
                 while (cellIterator.hasNext())
                 {
                	 Cell cell = cellIterator.next();
                     switch (cell.getCellTypeEnum()) 
                     {
                     	case STRING:
                     		System.out.print(cell.getStringCellValue() + "\t");
                     		cellData[i++]=cell.getStringCellValue();
                     		break;
                     	case NUMERIC:
                     		System.out.print(cell.getNumericCellValue() + "\t");
                     		cellData[i++]=cell.getNumericCellValue();
                     		break;
                     	case BOOLEAN:
                     		System.out.print(cell.getBooleanCellValue() + "\t");
                     		cellData[i++]=cell.getBooleanCellValue();
                     		break;
                     }
                 }
                 excel_data.put(key,cellData);
                 key++;
                 System.out.println();
             }
             //Writing data to excel from an array
             XSSFSheet sheetWrite = workbook.createSheet("Write_Emp_Data");
             int rownum = 0;
             for(Integer k : excel_data.keySet())
             {
            	 Row row = sheetWrite.createRow(rownum++);
            	 Object [] cellData = excel_data.get(k);
            	 int cellnum = 0;
                 for (Object obj : cellData)
                 {
                    Cell cell = row.createCell(cellnum++);
                    if(obj instanceof String)
                         cell.setCellValue((String)obj);
                     else if(obj instanceof Integer)
                         cell.setCellValue((Integer)obj);
                     else if(obj instanceof Double)
                         cell.setCellValue((Double)obj);
                     else if(obj instanceof Boolean)
                         cell.setCellValue((Boolean)obj);
                 }
             }
             FileOutputStream out = new FileOutputStream(excel);
             workbook.write(out);
             out.close();
             workbook.close();
    	}
    	catch(Exception e)
    	{
    		e.printStackTrace();
    	}
    }
}
