import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.*;

/**
* The MergeExcel program implements an application that
* merge two excel those have same id of employee's 
* but other information like address, Name are different.
* @author  Md. Gofran Khan
* @version 1.0
* @since   2019-06-03 
*/
public class MergeExcel {

	static XSSFRow row; 
	static XSSFRow row1;
	public static void main(String[] args) {
		// TODO Auto-generated method stub
		try{
	      // Opne two excel file
              File file = new File("E:\\ExcelsForPOI\\Employee_name_id.xlsx");
	      File file1 = new File("E:\\ExcelsForPOI\\Employee_id_and_more.xlsx");
	      FileInputStream fIP = new FileInputStream(file);
	      FileInputStream fIP1 = new FileInputStream(file1);
	      
	      //Get the workbook instance for XLSX file 
	      XSSFWorkbook workbook = new XSSFWorkbook(fIP);
	      XSSFWorkbook workbook1 = new XSSFWorkbook(fIP1);
	      
	      //Get the spreadsheet instance
	      XSSFSheet spreadsheet = workbook.getSheetAt(0);
	      Iterator < Row >  rowIterator = spreadsheet.iterator();
	      
	      ArrayList<Integer> ids = new ArrayList<Integer>(20);
	      
	      //Collect all id's from excel sheet one 
	      //and put in on id's ArrayList
	      while (rowIterator.hasNext()) {
	         row = (XSSFRow) rowIterator.next();
	         Iterator < Cell >  cellIterator = row.cellIterator();
	         while ( cellIterator.hasNext()) {
	            Cell cell = cellIterator.next();
	            switch (cell.getCellType()) {
	               case NUMERIC:
	                  ids.add((int)cell.getNumericCellValue());
	                  break;
	               
	               case STRING:
	                  break;
	            	}
	             }
	         }
	         boolean willPrint = false;
	         int id = 0;
	         String block = "";
	         String district = "";
	         int salary = 0;
	         ArrayList<ExcelData> excelData = new ArrayList<ExcelData>();
	         int index = 0;
		
		 //Find the identical id from others excel file
		 //and add them in ExcelData ArrayList 
	         for(int i = 0; i < 2; i++){
	         XSSFSheet spreadsheet1 = workbook1.getSheetAt(i);
	         Iterator < Row >  rowIterator1 = spreadsheet1.iterator();
	         while (rowIterator1.hasNext()) {
		         row1 = (XSSFRow) rowIterator1.next();
		         Iterator < Cell >  cellIterator1 = row1.cellIterator();
		         while ( cellIterator1.hasNext()) {
		            Cell cell1 = cellIterator1.next();
		            switch (cell1.getCellType()) {
		               case NUMERIC:
		                  for (int id1 : ids) {
					if(id1 == (int)cell1.getNumericCellValue())
					willPrint = true;
				  }
		                  if(willPrint){
		                	  switch(cell1.getColumnIndex()){
			            	  case 0:
			            		  id = (int)cell1.getNumericCellValue();
			            		  break;
			            	  case 3:
			            		  salary = ((int)cell1.getNumericCellValue());
			            		  break;
			            	  }
		                  }
		                  break;
		               
		               case STRING:
		            	  if(willPrint)
		            	  switch(cell1.getColumnIndex()){
		            	  case 1:
		            		  block = cell1.getStringCellValue();
		            		  break;
		            	  case 2:
		            		  district = cell1.getStringCellValue();
		            		  break;
		            	  }
		            }
		         }
		         if(willPrint){
		        	 excelData.add(new ExcelData(id, block, district, salary));
		         }
		         willPrint = false;
		         index++;
		         
		      }
	         }
	        for(int i = 0; i < excelData.size(); i++ ){
	        	ExcelData values = excelData.get(i);
	        	System.out.println(values.getId()+" "+ values.getBlock() +" "+ values.getDistrict() +" " + values.getSalary());
	        }
	      fIP.close();
	      fIP1.close();
		}catch(Exception e){
			e.printStackTrace();
		}
	}
}

/**
* The ExcelData program implements an application that
* create a property employee like id, block, district, salary etc 
* @author  Md. Gofran Khan
* @version 1.0
* @since   2019-06-03 
*/
class ExcelData{
	private int id;
	private String block;
	private String district;
	private int salary;

	ExcelData(int id, String block, String district, int salary){
		this.id = id;
		this.block = block;
		this.district = district;
		this.salary = salary;
	}
	
	public int getId() {
		return id;
	}
	public void setId(int id) {
		this.id = id;
	}
	public String getBlock() {
		return block;
	}
	public void setBlock(String block) {
		this.block = block;
	}
	public String getDistrict() {
		return district;
	}
	public void setDistrict(String district) {
		this.district = district;
	}
	public int getSalary() {
		return salary;
	}
	public void setSalary(int salary) {
		this.salary = salary;
	}
}
