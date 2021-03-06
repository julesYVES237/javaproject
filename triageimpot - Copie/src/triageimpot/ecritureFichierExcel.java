package triageimpot;

import java.io.File;
import java.io.IOException;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;


public class ecritureFichierExcel {
	
	public static void main(String[] args) {
		WritableWorkbook workbook = null;
		try {
			/* On cr�� un nouveau worbook et on l'ouvre en �criture */
			workbook = Workbook.createWorkbook(new File("C:\\Users\\user\\Desktop\\developpement\\eclipse\\WORKSPACE\\triageimpot\\src\\triageimpot\\recep.xls"));
			
			/* On cr�� une nouvelle feuille (test en position 0) et on l'ouvre en �criture */
			WritableSheet sheet = workbook.createSheet("test", 0); 
			
			/* Creation d'un champ au format texte */
			Label label = new Label(0, 0, "position A1");
			sheet.addCell(label);

			/* Creation d'un champ au format numerique */
			Number number = new Number(0, 1, 3.1459);
			sheet.addCell(number); 
			
			/* On ecrit le classeur */
			workbook.write(); 
			
		} 
		catch (IOException e) {
			e.printStackTrace();
		} 
		catch (RowsExceededException e) {
			e.printStackTrace();
		}
		catch (WriteException e) {
			e.printStackTrace();
		}
		finally {
			if(workbook!=null){
				/* On ferme le worbook pour lib�rer la m�moire */
				try {
					workbook.close();
				} 
				catch (WriteException e) {
					e.printStackTrace();
				} 
				catch (IOException e) {
					e.printStackTrace();
				} 
			}
		}

	}

}
