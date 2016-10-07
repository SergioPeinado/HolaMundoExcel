/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package holamundoexcel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author matinal
 */
public class LecturaExcel {
      public static void main(String[] args) {
          XSSFWorkbook wb;
          FileInputStream fis;
          Sheet sh;
          Row row;
          Cell cell;
          
          
          
	   try {
                fis = new FileInputStream("lecturaExcel.xlsx");
	        wb = new XSSFWorkbook(fis);
                //Primer for para las hojas
                for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                    sh = wb.getSheetAt(i);
                    System.out.print("\n##CARGADA HOJA = "+sh.getSheetName());
                    //Segundo for para las filas de las hojas
                    for (int j = 0; j < sh.getLastRowNum(); j++) {
                        row=sh.getRow(j);
                        System.out.print("\n");
                        //Tercer for para las celdas
                        for (int k = 0; k < row.getLastCellNum(); k++) {
                            cell = row.getCell(k);
                            System.out.print(cell.getAddress().toString()+" = "+cell+" ");
                        }
                    }
               }
                
                fis.close();
	   } catch (IOException ex) {
              Logger.getLogger(LecturaExcel.class.getName()).log(Level.SEVERE, null, ex);
          } finally {
	        
	   }
    }
}
