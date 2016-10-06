/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package holamundoexcel;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

/**
 *
 * @author matinal
 */
public class HolaMundoExcel {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
       
        SXSSFWorkbook wb = new SXSSFWorkbook(21); 
        Sheet sh = wb.createSheet("Hola Mundo");
        
        for (int i = 0; i < 26; i++) {
            Row row = sh.createRow(i);
            for (int j = 0; j < 25; j++) {
                Cell cell = row.createCell(j);
                cell.setCellValue((char) ('A'+j)+" "+(i+1));
            }
        }
        try {
            FileOutputStream out = new FileOutputStream("holaMundoExcel.xlsx");
            wb.write(out);
            out.close();
            
        } catch (IOException ex) {
           // Logger.getLogger(HolaMundoExcel.class.getName()).log(Level.SEVERE, null, ex);
            System.out.println("ERROR al crear el archivo: "+ex.getLocalizedMessage());
        }finally{
            wb.dispose();
        }
    }
    
}
