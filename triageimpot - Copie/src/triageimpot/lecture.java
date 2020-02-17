package triageimpot;
import java.io.File;


import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import triageimpot.lecture.classA;



public class lecture {
	 private static final String EXCEL_FILE_LOCATION2 ="C:\\Users\\user\\Desktop\\developpement\\eclipse\\WORKSPACE\\triageimpot\\DSF MICROSOFT_Normal_DGIFORMAT_VERROUILLEVF.xlsx";
		
	 static class classA{
	    	public  Date varA;
	    	public  String varB;
	    	public   String cell3;
	    	
	    }
	
		   
	 
	
	public static void main(String[] args) throws IOException {
		classA a = new classA();
		
		   WritableWorkbook myFirstWbook = null;
		  
		
		File excelFile =new File("C:\\Users\\user\\Desktop\\developpement\\eclipse\\WORKSPACE\\verifexcel\\SYSCOHADA DSF 2018 PIZZAROTTI SPA.xlsx");
		FileInputStream fils =new FileInputStream(excelFile);
		
		//maintenant on va creer un objet  xssfworkbook pour notre fichier  xlsx excel 
		
		XSSFWorkbook workbook = new XSSFWorkbook(fils);
		
		//nous nous positionnons sur la premiere feuille
		
		
		XSSFSheet sheet = workbook.getSheetAt(10);
		CreationHelper creationHelper1 = workbook.getCreationHelper();
		CellStyle style = workbook.createCellStyle();
		style.setDataFormat(creationHelper1.createDataFormat().getFormat(
				"dd-mm-yyyy"));
		
		Cell  cell = sheet.getRow(11).getCell(5);
		
		Cell cell3 = sheet.getRow(15).getCell(3);
		
		System.out.println(cell3);
	
		
		 cell.setCellStyle(style);
		   
		   
		System.out.println(cell);
		
		System.out.println(cell.getDateCellValue());
			
			
			a.varA = cell.getDateCellValue();
		
			
		System.out.println(a.varA);
		
		
		
	
	
    
    	 File FILE_NAME =new File(EXCEL_FILE_LOCATION2) ;

    	 FileInputStream fils1 =new FileInputStream(FILE_NAME);
        XSSFWorkbook workbook1 = new XSSFWorkbook(fils1);
        XSSFSheet sheet1 = workbook1.getSheetAt(4);
        CreationHelper creationHelper = workbook1.getCreationHelper();
     System.out.println(sheet1);
 
        
        Row row1 = sheet1.getRow(10);
       Cell cell2 = row1.createCell(5);
       CellStyle style1 = workbook1.createCellStyle();
       style1.setDataFormat(creationHelper.createDataFormat().getFormat(
				"dd-mm-yyyy"));
		cell2.setCellStyle(style1);
		SimpleDateFormat dt = new SimpleDateFormat("dd-mm-yyyy ");
		String date2 = dt.format(a.varA);
		//cell2 = date arrete effectif des compte (fr1)
       
       cell2.setCellValue( a.varA);
       System.out.println(cell2);
       
       Row row4 =sheet1.getRow(14);
       Cell cell4 =row4.createCell(4);
       cell4.setCellValue(a.cell3);
       System.out.println(cell4);
       
      
        try {
            FileOutputStream outputStream = new FileOutputStream(FILE_NAME);
            workbook1.write(outputStream);
            workbook1.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        System.out.println("Done");
    
}}

