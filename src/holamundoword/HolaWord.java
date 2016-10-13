/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package holamundoword;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;

/**
 *
 * @author matinal
 */
public class HolaWord {
    private static XWPFDocument doc=null;
    
    private static void testDoc(XWPFDocument doc){
                          
        for (int i = 0; i < 10; i++) {
             XWPFParagraph parr = doc.createParagraph();
            XWPFRun run = parr.createRun();
            run.setText(i+" : Hola Mundo Word.\n");
        }
        XWPFTable tabla = doc.createTable(12, 8);
        for (int i = 0; i < 12; i++) {
            for (int j = 0; j < 8; j++) {
                tabla.getRow(i).getCell(j).setText(i+":"+j);
            }
        }
        
        XWPFParagraph p = doc.createParagraph();
        XWPFRun r  = p.createRun();
        try{
        r.addPicture(new FileInputStream("CN_Logo.png"), XWPFDocument.PICTURE_TYPE_PNG, "CN_Logo.png", Units.toEMU(200), Units.toEMU(200));
       } catch(IOException | InvalidFormatException ex){
            System.out.println("ERROR: "+ex.getLocalizedMessage());
       }
        
        
    }
    
    public static void main(String[] args) {
        doc = new XWPFDocument();
        
        FileOutputStream out;
        try {
            out = new FileOutputStream(new File("holaWord.docx"));
            testDoc(doc);
            doc.write(out);
            out.close();
        } catch (IOException ex) {
            System.out.println("ERROR: "+ex.getLocalizedMessage());
        }
        
    }
}
