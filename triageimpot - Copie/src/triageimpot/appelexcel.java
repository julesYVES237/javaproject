package triageimpot;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.WritableSheet;

import java.io.File;
import java.io.IOException;

import jxl.write.*;






public class appelexcel {
	

    private static final String EXCEL_FILE_LOCATION = "C:\\Users\\user\\Desktop\\reception\\test\\Rattrapage Exars2024-1.xls";
    private static final String EXCEL_FILE_LOCATION2 = "C:\\Users\\user\\Desktop\\developpement\\eclipse\\WORKSPACE\\triageimpot\\src\\triageimpot\\recep.xls";

    static class classA{
    	public static String varA;
    	public static String varB;
    }
    public static void main(String[] args) {
    	
    	classA a = new classA();

        Workbook workbook = null;
        WritableWorkbook myFirstWbook = null;
      
        try {

            workbook = Workbook.getWorkbook(new File(EXCEL_FILE_LOCATION));

            Sheet sheet = workbook.getSheet(1);
            Cell cell1 = sheet.getCell(0, 0);
            System.out.print(cell1.getContents() + ":");    // Test Count + :
            Cell cell2 = sheet.getCell(0, 1);
            System.out.println(cell2.getContents());        // 1

            Cell cell3 = sheet.getCell(1, 0);
            System.out.print(cell3.getContents() + ":");    // Result + :
            Cell cell4 = sheet.getCell(1, 1);
            System.out.println(cell4.getContents());        // Passed

            System.out.print(cell1.getContents() + ":");    // Test Count + :
            cell2 = sheet.getCell(0, 2);
            System.out.println(cell2.getContents());        // 2

            System.out.print(cell3.getContents() + ":");    // Result + :
            cell4 = sheet.getCell(1, 2);
            System.out.println(cell4.getContents());        // Passed 2
            
            
            String contenuA1= cell1.getContents();
            String contenuA2= cell2.getContents();
            
            a.varA = cell1.getContents();
            a.varB = cell3.getContents();
            

           
            
            

        } catch (IOException e) {
            e.printStackTrace();
        } catch (BiffException e) {
            e.printStackTrace();
        } finally {

            if (workbook != null) {
                workbook.close();
            }

        }
        
        try {

            myFirstWbook = Workbook.createWorkbook(new File(EXCEL_FILE_LOCATION2));

            // create an Excel sheet
            WritableSheet excelSheet = myFirstWbook.createSheet("Sheet 1", 0);

            // add something into the Excel sheet
            Label label = new Label(0, 0, a.varB);
            excelSheet.addCell(label);

            Number number = new Number(0, 1, 1);
            excelSheet.addCell(number);

            label = new Label(1, 0, "Result");
            excelSheet.addCell(label);

            label = new Label(1, 1, a.varB);
            excelSheet.addCell(label);

            number = new Number(0, 2, 2);
            excelSheet.addCell(number);

            label = new Label(1, 2, a.varA);
            excelSheet.addCell(label);

            myFirstWbook.write();


        } catch (IOException e) {
            e.printStackTrace();
        } catch (WriteException e) {
            e.printStackTrace();
        } finally {

            if (myFirstWbook != null) {
                try {
                    myFirstWbook.close();
                } catch (IOException e) {
                    e.printStackTrace();
                } catch (WriteException e) {
                    e.printStackTrace();
                }
            }


        }


    }



	}

		
